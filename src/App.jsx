import { useEffect, useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import './App.css'

const DB_NAME = 'excel-editor-db'
const DB_VERSION = 1
const STORE_NAME = 'excel-datasets'

const openExcelDb = () =>
  new Promise((resolve, reject) => {
    const request = window.indexedDB.open(DB_NAME, DB_VERSION)

    request.onupgradeneeded = () => {
      const database = request.result

      if (!database.objectStoreNames.contains(STORE_NAME)) {
        const store = database.createObjectStore(STORE_NAME, {
          keyPath: 'id',
          autoIncrement: true,
        })
        store.createIndex('createdAt', 'createdAt')
      }
    }

    request.onsuccess = () => resolve(request.result)
    request.onerror = () => reject(request.error)
  })

const saveDatasetToDb = async (payload) => {
  const database = await openExcelDb()

  return new Promise((resolve, reject) => {
    const transaction = database.transaction(STORE_NAME, 'readwrite')
    const store = transaction.objectStore(STORE_NAME)
    const request = store.add(payload)

    request.onsuccess = () => resolve(request.result)
    request.onerror = () => reject(request.error)
    transaction.oncomplete = () => database.close()
    transaction.onerror = () => database.close()
  })
}

const loadLatestDatasetFromDb = async () => {
  const database = await openExcelDb()

  return new Promise((resolve, reject) => {
    const transaction = database.transaction(STORE_NAME, 'readonly')
    const store = transaction.objectStore(STORE_NAME)
    const request = store.getAll()

    request.onsuccess = () => {
      const records = request.result || []

      if (!records.length) {
        resolve(null)
        return
      }

      const latestRecord = records
        .slice()
        .sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt))[0]

      resolve(latestRecord)
    }

    request.onerror = () => reject(request.error)
    transaction.oncomplete = () => database.close()
    transaction.onerror = () => database.close()
  })
}

const loadAllDatasetsFromDb = async () => {
  const database = await openExcelDb()

  return new Promise((resolve, reject) => {
    const transaction = database.transaction(STORE_NAME, 'readonly')
    const store = transaction.objectStore(STORE_NAME)
    const request = store.getAll()

    request.onsuccess = () => {
      const records = request.result || []
      const sortedRecords = records
        .slice()
        .sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt))

      resolve(sortedRecords)
    }

    request.onerror = () => reject(request.error)
    transaction.oncomplete = () => database.close()
    transaction.onerror = () => database.close()
  })
}

const deleteDatasetFromDb = async (datasetId) => {
  const database = await openExcelDb()

  return new Promise((resolve, reject) => {
    const transaction = database.transaction(STORE_NAME, 'readwrite')
    const store = transaction.objectStore(STORE_NAME)
    const request = store.delete(datasetId)

    request.onsuccess = () => resolve()
    request.onerror = () => reject(request.error)
    transaction.oncomplete = () => database.close()
    transaction.onerror = () => database.close()
  })
}

function App() {
  const [selectedFile, setSelectedFile] = useState(null)
  const [message, setMessage] = useState('No file selected yet.')
  const [status, setStatus] = useState('idle')
  const [tableData, setTableData] = useState([])
  const [headers, setHeaders] = useState([])
  const [isPreviewVisible, setIsPreviewVisible] = useState(false)
  const [cellErrors, setCellErrors] = useState({})
  const [columnBSortDirection, setColumnBSortDirection] = useState('asc')
  const [savedDatasets, setSavedDatasets] = useState([])
  const [datasetSearchTerm, setDatasetSearchTerm] = useState('')

  const acceptedTypes = useMemo(
    () => [
      'application/vnd.ms-excel',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'text/csv',
      'application/csv',
    ],
    [],
  )

  const filteredDatasets = useMemo(() => {
    const normalizedSearch = datasetSearchTerm.trim().toLowerCase()

    if (!normalizedSearch) {
      return savedDatasets
    }

    return savedDatasets.filter((dataset) =>
      String(dataset.fileName || 'excel-data')
        .toLowerCase()
        .includes(normalizedSearch),
    )
  }, [savedDatasets, datasetSearchTerm])

  const isExcelFile = (file) => {
    const lowerName = file.name.toLowerCase()
    const hasValidExtension =
      lowerName.endsWith('.xls') || lowerName.endsWith('.xlsx') || lowerName.endsWith('.csv')

    return hasValidExtension || acceptedTypes.includes(file.type)
  }

  const getCellError = (rowIndex, cellIndex, value, rowValues = []) => {
    const excelRowNumber = rowIndex + 2
    const rawValue = String(value ?? '')
    const trimmedValue = String(value ?? '').trim()

    if (excelRowNumber < 6) {
      return ''
    }

    if (cellIndex === 1) {
      if (trimmedValue.length === 0) {
        return 'Column B from row 6 is mandatory.'
      }

      if (trimmedValue.length > 6) {
        return 'Column B from row 6 must be 6 characters or fewer.'
      }

      if (!/^[A-Z]+$/.test(trimmedValue)) {
        return 'Column B from row 6 must contain only uppercase letters (A-Z).'
      }
    }

    if (cellIndex === 2) {
      const integerFormat = /^-?\d{1,2}$/

      if (trimmedValue.length === 0) {
        return 'Column C from row 6 is mandatory.'
      }

      if (!integerFormat.test(trimmedValue)) {
        return 'Column C from row 6 must be a valid int(2) value.'
      }
    }

    if (cellIndex === 3) {
      const integerFormat = /^-?\d{1,2}$/

      if (trimmedValue.length === 0) {
        return 'Column D from row 6 is mandatory.'
      }

      if (!integerFormat.test(trimmedValue)) {
        return 'Column D from row 6 must be a valid int(2) value.'
      }
    }

    if (cellIndex === 4) {
      const integerFormat = /^-?\d{1,2}$/

      if (trimmedValue.length === 0) {
        return 'Column E from row 6 is mandatory.'
      }

      if (!integerFormat.test(trimmedValue)) {
        return 'Column E from row 6 must be a valid int(2) value.'
      }
    }

    if (cellIndex === 5) {
      if (trimmedValue.length === 0) {
        return 'Column F from row 6 is mandatory.'
      }

      if (trimmedValue.length > 48) {
        return 'Column F from row 6 must be 48 characters or fewer.'
      }

      if (trimmedValue.length > 0 && !/^[A-Z0-9 !"#$%&'()*+,\-./:;<=>?@[\\\]^_`{|}~]+$/.test(trimmedValue)) {
        return 'Column F from row 6 must contain uppercase letters (A-Z), numbers (0-9), spaces, or special characters only.'
      }
    }

    if (cellIndex === 6) {
      if (trimmedValue.length > 4) {
        return 'Column G from row 6 must be 4 characters or fewer.'
      }

      if (trimmedValue.length > 0 && !/^[A-Z]+$/.test(trimmedValue)) {
        return 'Column G from row 6 must contain only uppercase letters (A-Z).'
      }
    }

    if (cellIndex === 7) {
      if (trimmedValue.length === 0) {
        return 'Column H from row 6 is mandatory.'
      }

      if (trimmedValue.length > 4) {
        return 'Column H from row 6 must be 4 characters or fewer.'
      }

      if (trimmedValue.length > 0 && !/^[A-Z]+$/.test(trimmedValue)) {
        return 'Column H from row 6 must contain only uppercase letters (A-Z).'
      }
    }

    if (cellIndex === 8) {
      if (trimmedValue.length === 0) {
        return 'Column I from row 6 is mandatory.'
      }

      if (trimmedValue.length > 4) {
        return 'Column I from row 6 must be 4 characters or fewer.'
      }

      if (trimmedValue.length > 0 && !/^[A-Z]+$/.test(trimmedValue)) {
        return 'Column I from row 6 must contain only uppercase letters (A-Z).'
      }
    }

    if (cellIndex === 9) {
      if (trimmedValue.length === 0) {
        return 'Column J from row 6 is mandatory.'
      }

      if (trimmedValue.length > 4) {
        return 'Column J from row 6 must be 4 characters or fewer.'
      }

      if (trimmedValue.length > 0 && !/^[A-Z]+$/.test(trimmedValue)) {
        return 'Column J from row 6 must contain only uppercase letters (A-Z).'
      }
    }

    if (cellIndex === 10) {
      if (trimmedValue.length === 0) {
        return 'Column K from row 6 is mandatory.'
      }

      if (trimmedValue.length > 4) {
        return 'Column K from row 6 must be 4 characters or fewer.'
      }

      if (trimmedValue.length > 0 && !/^[A-Z]+$/.test(trimmedValue)) {
        return 'Column K from row 6 must contain only uppercase letters (A-Z).'
      }
    }

    if (cellIndex === 11) {
      const decimalFormat = /^-?\d{1,4}(\.\d{1,3})?$/

      if (trimmedValue.length === 0) {
        return 'Column L from row 6 is mandatory.'
      }

      if (!decimalFormat.test(trimmedValue)) {
        return 'Column L from row 6 must be a valid decimal(7,3) value.'
      }
    }

    if (cellIndex === 17) {
      const decimalFormat = /^-?\d{1,4}(\.\d{1,3})?$/

      if (trimmedValue.length === 0) {
        return 'Column R from row 6 is mandatory.'
      }

      if (!decimalFormat.test(trimmedValue)) {
        return 'Column R from row 6 must be a valid decimal(7,3) value.'
      }
    }

    if (cellIndex === 15) {
      const decimalFormat = /^-?\d{1,4}(\.\d{1,3})?$/

      if (trimmedValue.length === 0) {
        return 'Column P from row 6 is mandatory.'
      }

      if (!decimalFormat.test(trimmedValue)) {
        return 'Column P from row 6 must be a valid decimal(7,3) value.'
      }
    }

    if (cellIndex === 13) {
      const decimalFormat = /^-?\d{1,4}(\.\d{1,3})?$/

      if (trimmedValue.length === 0) {
        return 'Column N from row 6 is mandatory.'
      }

      if (!decimalFormat.test(trimmedValue)) {
        return 'Column N from row 6 must be a valid decimal(7,3) value.'
      }
    }

    if (cellIndex === 12) {
      if (trimmedValue.length === 0) {
        return 'Column M from row 6 is mandatory.'
      }

      const normalizedFlag = trimmedValue.toUpperCase()

      if (!(trimmedValue.length === 1 && (normalizedFlag === 'Y' || normalizedFlag === 'N'))) {
        return "Column M from row 6 must be exactly 1 character: 'Y' or 'N'."
      }
    }

    if (cellIndex === 14) {
      if (trimmedValue.length === 0) {
        return 'Column O from row 6 is mandatory.'
      }

      const normalizedFlag = trimmedValue.toUpperCase()

      if (!(trimmedValue.length === 1 && (normalizedFlag === 'Y' || normalizedFlag === 'N'))) {
        return "Column O from row 6 must be exactly 1 character: 'Y' or 'N'."
      }
    }

    if (cellIndex === 16) {
      if (trimmedValue.length === 0) {
        return 'Column Q from row 6 is mandatory.'
      }

      const normalizedFlag = rawValue.toUpperCase()
      const allowedValues = ['Y', 'N']

      if (!(rawValue.length === 1 && allowedValues.includes(normalizedFlag))) {
        return "Column Q from row 6 must be exactly 1 character: 'Y' or 'N'."
      }
    }

    if (cellIndex === 18) {
      if (trimmedValue.length === 0) {
        return 'Column S from row 6 is mandatory.'
      }

      const normalizedFlag = trimmedValue.toUpperCase()

      if (!(trimmedValue.length === 1 && (normalizedFlag === 'Y' || normalizedFlag === 'N'))) {
        return "Column S from row 6 must be exactly 1 character: 'Y' or 'N'."
      }
    }

    if (cellIndex === 19) {
      if (trimmedValue.length === 0) {
        return 'Column T from row 6 is mandatory.'
      }

      if (trimmedValue.length > 4) {
        return 'Column T from row 6 must be 4 characters or fewer.'
      }

      if (!/^[A-Z]+$/.test(trimmedValue)) {
        return 'Column T from row 6 must contain only uppercase letters (A-Z).'
      }
    }

    if (cellIndex === 20) {
      if (trimmedValue.length === 0) {
        return 'Column U from row 6 is mandatory.'
      }

      if (trimmedValue.length > 4) {
        return 'Column U from row 6 must be 4 characters or fewer.'
      }

      if (!/^[A-Z]+$/.test(trimmedValue)) {
        return 'Column U from row 6 must contain only uppercase letters (A-Z).'
      }
    }

    if (cellIndex === 21) {
      if (trimmedValue.length === 0) {
        return 'Column V from row 6 is mandatory.'
      }

      if (trimmedValue.length > 4) {
        return 'Column V from row 6 must be 4 characters or fewer.'
      }

      if (!/^[A-Z]+$/.test(trimmedValue)) {
        return 'Column V from row 6 must contain only uppercase letters (A-Z).'
      }
    }

    if (cellIndex === 22 && trimmedValue.length > 15) {
      return 'Column W from row 6 must be 15 characters or fewer.'
    }

    if (cellIndex === 23 && trimmedValue.length > 15) {
      return 'Column X from row 6 must be 15 characters or fewer.'
    }

    if (cellIndex === 24) {
      const normalizedFlag = rawValue.toUpperCase()
      const allowedValues = ['S', 'A', 'G', 'F', 'M', 'R']

      if (rawValue.length > 0 && !(rawValue.length === 1 && allowedValues.includes(normalizedFlag))) {
        return "Column Y from row 6 must be exactly 1 character: 'S', 'A', 'G', 'F', 'M', or 'R'."
      }
    }

    if (cellIndex === 25 && trimmedValue.length > 24) {
      return 'Column Z from row 6 must be 24 characters or fewer.'
    }

    if (cellIndex === 26) {
      if (trimmedValue.length > 1) {
        return 'Column AA from row 6 must be 1 character or fewer.'
      }

      if (trimmedValue.length === 1 && trimmedValue !== 'Y' && trimmedValue !== 'N') {
        return "Column AA from row 6 must be 'Y' or 'N' in uppercase when provided."
      }
    }

    if (cellIndex === 27) {
      if (trimmedValue.length > 48) {
        return 'Column AB from row 6 must be 48 characters or fewer.'
      }

      if (trimmedValue.length > 0 && !/^[A-Z0-9 !"#$%&'()*+,\-./:;<=>?@[\\\]^_`{|}~]+$/.test(trimmedValue)) {
        return 'Column AB from row 6 must contain uppercase letters (A-Z), numbers (0-9), spaces, or special characters only.'
      }
    }

    if (cellIndex === 28) {
      if (trimmedValue.length > 48) {
        return 'Column AC from row 6 must be 48 characters or fewer.'
      }

      if (trimmedValue.length > 0 && !/^[A-Z0-9 !"#$%&'()*+,\-./:;<=>?@[\\\]^_`{|}~]+$/.test(trimmedValue)) {
        return 'Column AC from row 6 must contain uppercase letters (A-Z), numbers (0-9), spaces, or special characters only.'
      }
    }

    if (cellIndex === 29) {
      if (trimmedValue.length > 48) {
        return 'Column AD from row 6 must be 48 characters or fewer.'
      }

      if (trimmedValue.length > 0 && !/^[A-Z0-9 !"#$%&'()*+,\-./:;<=>?@[\\\]^_`{|}~]+$/.test(trimmedValue)) {
        return 'Column AD from row 6 must contain uppercase letters (A-Z), numbers (0-9), spaces, or special characters only.'
      }
    }

    if (cellIndex === 30 && trimmedValue.length > 8) {
      return 'Column AE from row 6 must be 8 characters or fewer.'
    }

    if (cellIndex === 31 && trimmedValue.length > 8) {
      return 'Column AF from row 6 must be 8 characters or fewer.'
    }

    if (cellIndex === 32 && trimmedValue.length > 2) {
      return 'Column AG from row 6 must be 2 characters or fewer.'
    }

    if (cellIndex === 33) {
      if (trimmedValue.length > 10) {
        return 'Column AH from row 6 must be 10 characters or fewer.'
      }

      if (trimmedValue.length > 0 && !/^[A-Z]+$/.test(trimmedValue)) {
        return 'Column AH from row 6 must contain only uppercase letters (A-Z).'
      }
    }

    if (cellIndex === 34) {
      if (trimmedValue.length > 10) {
        return 'Column AI from row 6 must be 10 characters or fewer.'
      }
    }

    if (cellIndex === 35) {
      if (trimmedValue.length > 20) {
        return 'Column AJ from row 6 must be 20 characters or fewer.'
      }

      if (trimmedValue.length > 0 && !/^[A-Z]+$/.test(trimmedValue)) {
        return 'Column AJ from row 6 must contain only uppercase letters (A-Z).'
      }
    }

    if (cellIndex === 36) {
      if (trimmedValue.length > 20) {
        return 'Column AK from row 6 must be 20 characters or fewer.'
      }

      if (trimmedValue.length > 0 && !/^[A-Z]+$/.test(trimmedValue)) {
        return 'Column AK from row 6 must contain only uppercase letters (A-Z).'
      }
    }

    if (cellIndex === 37) {
      if (trimmedValue.length > 20) {
        return 'Column AL from row 6 must be 20 characters or fewer.'
      }

      if (trimmedValue.length > 0 && !/^[A-Z]+$/.test(trimmedValue)) {
        return 'Column AL from row 6 must contain only uppercase letters (A-Z).'
      }
    }

    if (cellIndex === 38) {
      if (trimmedValue.length > 1) {
        return 'Column AM from row 6 must be 1 character or fewer.'
      }

      if (trimmedValue.length > 0 && !/^[A-Z]+$/.test(trimmedValue)) {
        return 'Column AM from row 6 must contain only uppercase letters (A-Z).'
      }
    }

    if (cellIndex === 39) {
      if (trimmedValue.length > 1) {
        return 'Column AN from row 6 must be 1 character or fewer.'
      }

      if (trimmedValue.length > 0 && !/^[A-Z]+$/.test(trimmedValue)) {
        return 'Column AN from row 6 must contain only uppercase letters (A-Z).'
      }
    }

    if (cellIndex === 40) {
      const decimalFormat = /^-?\d{1,7}(\.\d{1,6})?$/

      if (trimmedValue.length > 0 && !decimalFormat.test(trimmedValue)) {
        return 'Column AO from row 6 must be a valid decimal(13,6) value.'
      }
    }

    if (cellIndex === 41) {
      const decimalFormat = /^-?\d{1,5}(\.\d{1,2})?$/

      if (trimmedValue.length > 0 && !decimalFormat.test(trimmedValue)) {
        return 'Column AP from row 6 must be a valid decimal(7,2) value.'
      }
    }

    if (cellIndex === 42) {
      if (trimmedValue.length > 3) {
        return 'Column AQ from row 6 must be 3 characters or fewer.'
      }

      if (trimmedValue.length > 0 && !/^[A-Z]+$/.test(trimmedValue)) {
        return 'Column AQ from row 6 must contain only uppercase letters (A-Z).'
      }
    }

    if (cellIndex === 43) {
      if (trimmedValue.length > 3) {
        return 'Column AR from row 6 must be 3 characters or fewer.'
      }

      if (trimmedValue.length > 0 && !/^[A-Z]+$/.test(trimmedValue)) {
        return 'Column AR from row 6 must contain only uppercase letters (A-Z).'
      }
    }

    if (cellIndex === 44) {
      if (trimmedValue.length > 5) {
        return 'Column AS from row 6 must be 5 characters or fewer.'
      }

      if (trimmedValue.length > 0 && !/^[A-Z]+$/.test(trimmedValue)) {
        return 'Column AS from row 6 must contain only uppercase letters (A-Z).'
      }
    }

    if (cellIndex === 45) {
      if (trimmedValue.length > 3) {
        return 'Column AT from row 6 must be 3 characters or fewer.'
      }

      if (trimmedValue.length > 0 && !/^[A-Z]+$/.test(trimmedValue)) {
        return 'Column AT from row 6 must contain only uppercase letters (A-Z).'
      }
    }

    if (cellIndex === 46) {
      if (trimmedValue.length > 3) {
        return 'Column AU from row 6 must be 3 characters or fewer.'
      }

      if (trimmedValue.length > 0 && !/^[A-Z]+$/.test(trimmedValue)) {
        return 'Column AU from row 6 must contain only uppercase letters (A-Z).'
      }
    }

    if (cellIndex === 47) {
      const decimalFormat = /^-?\d{1,10}(\.\d{1,3})?$/

      if (trimmedValue.length > 0 && !decimalFormat.test(trimmedValue)) {
        return 'Column AV from row 6 must be a valid decimal(13,3) value.'
      }
    }

    if (cellIndex === 48) {
      const decimalFormat = /^-?\d{1,10}(\.\d{1,3})?$/

      if (trimmedValue.length > 0 && !decimalFormat.test(trimmedValue)) {
        return 'Column AW from row 6 must be a valid decimal(13,3) value.'
      }
    }

    if (cellIndex === 49) {
      if (trimmedValue.length > 4) {
        return 'Column AX from row 6 must be 4 characters or fewer.'
      }

      if (trimmedValue.length > 0 && !/^[A-Z]+$/.test(trimmedValue)) {
        return 'Column AX from row 6 must contain only uppercase letters (A-Z).'
      }
    }

    if (cellIndex === 50) {
      if (trimmedValue.length > 4) {
        return 'Column AY from row 6 must be 4 characters or fewer.'
      }

      if (trimmedValue.length > 0 && !/^[A-Z]+$/.test(trimmedValue)) {
        return 'Column AY from row 6 must contain only uppercase letters (A-Z).'
      }
    }

    if (cellIndex === 51) {
      if (trimmedValue.length > 4) {
        return 'Column AZ from row 6 must be 4 characters or fewer.'
      }

      if (trimmedValue.length > 0 && !/^[A-Z]+$/.test(trimmedValue)) {
        return 'Column AZ from row 6 must contain only uppercase letters (A-Z).'
      }
    }

    if (cellIndex === 52) {
      if (trimmedValue.length > 1) {
        return 'Column BA from row 6 must be 1 character or fewer.'
      }

      if (trimmedValue.length > 0 && !/^[A-Z]+$/.test(trimmedValue)) {
        return 'Column BA from row 6 must contain only uppercase letters (A-Z).'
      }
    }

    if (cellIndex === 53) {
      const integerFormat = /^-?\d{1,4}$/

      if (!integerFormat.test(trimmedValue)) {
        return 'Column BB from row 6 must be a valid int(4) value.'
      }
    }

    if (cellIndex === 54) {
      if (trimmedValue.length > 1) {
        return 'Column BC from row 6 must be 1 character or fewer.'
      }

      if (trimmedValue.length > 0 && !/^[A-Z]+$/.test(trimmedValue)) {
        return 'Column BC from row 6 must contain only uppercase letters (A-Z).'
      }
    }

    if (cellIndex === 55) {
      const integerFormat = /^-?\d{1,10}$/

      if (!integerFormat.test(trimmedValue)) {
        return 'Column BD from row 6 must be a valid int(10) value.'
      }
    }

    if (cellIndex === 56) {
      if (trimmedValue.length > 1) {
        return 'Column BE from row 6 must be 1 character or fewer.'
      }

      if (trimmedValue.length > 0 && !/^[A-Z]+$/.test(trimmedValue)) {
        return 'Column BE from row 6 must contain only uppercase letters (A-Z).'
      }
    }

    if (cellIndex === 57) {
      if (trimmedValue.length > 1) {
        return 'Column BF from row 6 must be 1 character or fewer.'
      }

      if (trimmedValue.length > 0 && !/^[A-Z]+$/.test(trimmedValue)) {
        return 'Column BF from row 6 must contain only uppercase letters (A-Z).'
      }
    }

    if (cellIndex === 58) {
      if (trimmedValue.length > 200) {
        return 'Column BG from row 6 must be 200 characters or fewer.'
      }

      if (trimmedValue.length > 0 && !/^[A-Z]+$/.test(trimmedValue)) {
        return 'Column BG from row 6 must contain only uppercase letters (A-Z).'
      }
    }

    return ''
  }

  const normalizeCellValue = (value) => String(value ?? '').trim()

  const collectCellErrors = (rows) => {
    const nextErrors = {}

    rows.forEach((row, rowIndex) => {
      row.forEach((cell, cellIndex) => {
        const errorMessage = getCellError(rowIndex, cellIndex, cell, row)

        if (errorMessage) {
          nextErrors[`${rowIndex}-${cellIndex}`] = errorMessage
        }
      })
    })

    return nextErrors
  }

  const handleFileChange = (event) => {
    const file = event.target.files?.[0]

    if (!file) {
      setSelectedFile(null)
      setStatus('idle')
      setMessage('No file selected yet.')
      setHeaders([])
      setTableData([])
      setIsPreviewVisible(false)
      setCellErrors({})
      return
    }

    if (!isExcelFile(file)) {
      setSelectedFile(null)
      setStatus('error')
      setMessage('Please choose a valid file (.xls, .xlsx, or .csv).')
      setHeaders([])
      setTableData([])
      setIsPreviewVisible(false)
      setCellErrors({})
      return
    }

    setSelectedFile(file)
    setStatus('success')
    setMessage('File is ready to upload.')
    setHeaders([])
    setTableData([])
    setIsPreviewVisible(false)
    setCellErrors({})
  }

  const handleUpload = async () => {
    if (!selectedFile) {
      setStatus('error')
      setMessage('Choose a file before uploading.')
      return
    }

    try {
      const buffer = await selectedFile.arrayBuffer()
      const workbook = XLSX.read(buffer, { type: 'array' })
      const firstSheetName = workbook.SheetNames[0]

      if (!firstSheetName) {
        setStatus('error')
        setMessage('The selected Excel file has no sheets.')
        setHeaders([])
        setTableData([])
        setIsPreviewVisible(false)
        setCellErrors({})
        return
      }

      const sheet = workbook.Sheets[firstSheetName]
      const rows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        blankrows: false,
        defval: '',
      })

      if (!rows.length) {
        setStatus('error')
        setMessage('No data found in the first sheet.')
        setHeaders([])
        setTableData([])
        setIsPreviewVisible(false)
        setCellErrors({})
        return
      }

      const [headerRow, ...bodyRows] = rows
      const normalizedHeaders = headerRow.map((cell, index) =>
        String(cell).trim() || `Column ${index + 1}`,
      )

      const normalizedRows = bodyRows.map((row) =>
        normalizedHeaders.map((_, index) => normalizeCellValue(row[index])),
      )
      const validationErrors = collectCellErrors(normalizedRows)
      const errorCount = Object.keys(validationErrors).length

      setHeaders(normalizedHeaders)
      setTableData(normalizedRows)
      setCellErrors(validationErrors)
      setIsPreviewVisible(true)

      if (errorCount > 0) {
        setStatus('error')
        setMessage(`Validation found ${errorCount} error(s). Review highlighted cells in preview.`)
      } else {
        setStatus('success')
        setMessage(
          `Loaded ${bodyRows.length} rows from "${firstSheetName}" in ${selectedFile.name}.`,
        )
      }
    } catch {
      setStatus('error')
      setMessage('Unable to read this Excel file. Please try another file.')
      setHeaders([])
      setTableData([])
      setIsPreviewVisible(false)
      setCellErrors({})
    }
  }

  const handleCellChange = (rowIndex, cellIndex, value) => {
    const rawValue = String(value ?? '')

    setTableData((currentRows) =>
      {
        const nextRows = currentRows.map((row, currentRowIndex) => {
        if (currentRowIndex !== rowIndex) {
          return row
        }

        return row.map((cell, currentCellIndex) =>
          currentCellIndex === cellIndex ? rawValue : cell,
        )
        })

        setCellErrors(collectCellErrors(nextRows))
        return nextRows
      },
    )
  }

  const handleAddNewLine = () => {
    if (!headers.length) {
      setStatus('error')
      setMessage('Load Excel data before adding a new line.')
      return
    }

    const newRow = headers.map(() => '')

    setTableData((currentRows) => {
      const nextRows = [...currentRows, newRow]
      setCellErrors(collectCellErrors(nextRows))
      return nextRows
    })

    setStatus('success')
    setMessage('New line added to Excel Data Preview.')
  }

  const handleDeleteRow = (rowIndexToDelete) => {
    const firstDeletableRowIndex = 4

    if (rowIndexToDelete < firstDeletableRowIndex) {
      setStatus('error')
      setMessage('Rows above row 6 cannot be deleted.')
      return
    }

    setTableData((currentRows) => {
      if (rowIndexToDelete < 0 || rowIndexToDelete >= currentRows.length) {
        return currentRows
      }

      const nextRows = currentRows.filter((_, currentRowIndex) => currentRowIndex !== rowIndexToDelete)
      setCellErrors(collectCellErrors(nextRows))
      return nextRows
    })

    setStatus('success')
    setMessage(`Deleted row ${rowIndexToDelete + 1} from preview.`)
  }

  const handleSortColumnB = () => {
    if (!headers.length || headers.length < 2) {
      setStatus('error')
      setMessage('Column B is not available for sorting.')
      return
    }

    const nextDirection = columnBSortDirection === 'asc' ? 'desc' : 'asc'

    setTableData((currentRows) => {
      const sortStartIndex = 4

      if (currentRows.length <= sortStartIndex) {
        return currentRows
      }

      const fixedRows = currentRows.slice(0, sortStartIndex)
      const sortableRows = [...currentRows.slice(sortStartIndex)].sort((a, b) => {
        const aValue = String(a[1] ?? '').trim().toLowerCase()
        const bValue = String(b[1] ?? '').trim().toLowerCase()

        if (aValue === bValue) {
          return 0
        }

        if (nextDirection === 'asc') {
          return aValue.localeCompare(bValue)
        }

        return bValue.localeCompare(aValue)
      })

      const sortedRows = [...fixedRows, ...sortableRows]

      setCellErrors(collectCellErrors(sortedRows))
      return sortedRows
    })

    setColumnBSortDirection(nextDirection)
    setStatus('success')
    setMessage(
      `Sorted by Column B from row 6 (${nextDirection === 'asc' ? 'ascending' : 'descending'}).`,
    )
  }

  const handleConvertRow6ToUppercase = () => {
    if (!tableData.length) {
      setStatus('error')
      setMessage('No data available to convert.')
      return
    }

    setTableData((currentRows) => {
      const nextRows = [...currentRows]
      const row6Index = 4

      if (row6Index >= nextRows.length) {
        setStatus('error')
        setMessage('Row 6 does not exist in the data.')
        return nextRows
      }

      nextRows[row6Index] = nextRows[row6Index].map((cell) => String(cell ?? '').toUpperCase())
      setCellErrors(collectCellErrors(nextRows))
      return nextRows
    })

    setStatus('success')
    setMessage('Converted row 6 text to uppercase.')
  }

  const handleSortByErrorRow = () => {
    if (!headers.length) {
      setStatus('error')
      setMessage('Load Excel data before sorting by error rows.')
      return
    }

    setTableData((currentRows) => {
      const sortStartIndex = 4

      if (currentRows.length <= sortStartIndex) {
        return currentRows
      }

      const fixedRows = currentRows.slice(0, sortStartIndex)
      const sortableRows = currentRows
        .slice(sortStartIndex)
        .map((row, originalIndex) => ({
          row,
          originalIndex,
          hasError: row.some((cell, cellIndex) =>
            Boolean(getCellError(sortStartIndex + originalIndex, cellIndex, cell, row)),
          ),
        }))
        .sort((a, b) => {
          if (a.hasError === b.hasError) {
            return a.originalIndex - b.originalIndex
          }

          return a.hasError ? -1 : 1
        })
        .map((item) => item.row)

      const sortedRows = [...fixedRows, ...sortableRows]
      setCellErrors(collectCellErrors(sortedRows))
      return sortedRows
    })

    setStatus('success')
    setMessage('Sorted by error rows from row 6 (error rows first).')
  }

  const handleSaveExcel = () => {
    if (!headers.length) {
      setStatus('error')
      setMessage('Load Excel data before saving.')
      return
    }

    try {
      const rowsToExport = [headers, ...tableData]
      const exportSheet = XLSX.utils.aoa_to_sheet(rowsToExport)
      const exportBook = XLSX.utils.book_new()
      XLSX.utils.book_append_sheet(exportBook, exportSheet, 'EditedData')

      const baseName = selectedFile?.name
        ? selectedFile.name.replace(/\.[^.]+$/, '')
        : 'excel-data'
      const suggestedName = `${baseName}-edited.xlsx`
      const enteredName = window.prompt('Enter Excel file name to save:', suggestedName)

      if (enteredName === null) {
        setStatus('idle')
        setMessage('Excel save cancelled.')
        return
      }

      const trimmedName = enteredName.trim()

      if (!trimmedName) {
        setStatus('error')
        setMessage('Please enter a valid file name.')
        return
      }

      const safeName = trimmedName.replace(/[\\/:*?"<>|]/g, '_')
      const outputName = /\.(xlsx|xls)$/i.test(safeName) ? safeName : `${safeName}.xlsx`

      XLSX.writeFile(exportBook, outputName)
      setStatus('success')
      setMessage(`Saved edited data as ${outputName}`)
    } catch {
      setStatus('error')
      setMessage('Could not save Excel file. Please try again.')
    }
  }

  const handleSaveCsv = () => {
    if (!headers.length) {
      setStatus('error')
      setMessage('Load data before saving CSV.')
      return
    }

    try {
      const rowsToExport = [headers, ...tableData]
      const exportSheet = XLSX.utils.aoa_to_sheet(rowsToExport)
      const csvContent = XLSX.utils.sheet_to_csv(exportSheet)

      const baseName = selectedFile?.name
        ? selectedFile.name.replace(/\.[^.]+$/, '')
        : 'excel-data'
      const suggestedName = `${baseName}-edited.csv`
      const enteredName = window.prompt('Enter CSV file name to save:', suggestedName)

      if (enteredName === null) {
        setStatus('idle')
        setMessage('CSV save cancelled.')
        return
      }

      const trimmedName = enteredName.trim()

      if (!trimmedName) {
        setStatus('error')
        setMessage('Please enter a valid file name.')
        return
      }

      const safeName = trimmedName.replace(/[\\/:*?"<>|]/g, '_')
      const outputName = /\.csv$/i.test(safeName) ? safeName : `${safeName}.csv`

      const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' })
      const url = URL.createObjectURL(blob)
      const link = document.createElement('a')

      link.href = url
      link.download = outputName
      link.style.display = 'none'
      document.body.appendChild(link)
      link.click()
      document.body.removeChild(link)
      URL.revokeObjectURL(url)

      setStatus('success')
      setMessage(`Saved edited data as ${outputName}`)
    } catch {
      setStatus('error')
      setMessage('Could not save CSV file. Please try again.')
    }
  }

  const handleSaveToLocalDb = async () => {
    if (!headers.length) {
      setStatus('error')
      setMessage('Load Excel data before saving to local database.')
      return
    }

    try {
      await saveDatasetToDb({
        fileName: selectedFile?.name || 'excel-data',
        headers,
        rows: tableData,
        createdAt: new Date().toISOString(),
      })

      const allRecords = await loadAllDatasetsFromDb()
      setSavedDatasets(allRecords)

      setStatus('success')
      setMessage('Saved data to local database (IndexedDB).')
    } catch {
      setStatus('error')
      setMessage('Unable to save data in local database.')
    }
  }

  const handleLoadFromLocalDb = async () => {
    try {
      const record = await loadLatestDatasetFromDb()

      if (!record) {
        setStatus('error')
        setMessage('No saved dataset found in local database.')
        return
      }

      const normalizedLoadedRows = (record.rows || []).map((row) =>
        row.map((cell) => normalizeCellValue(cell)),
      )

      setHeaders(record.headers || [])
      setTableData(normalizedLoadedRows)
      setCellErrors(collectCellErrors(normalizedLoadedRows))
      setIsPreviewVisible(true)
      setStatus('success')
      setMessage(`Loaded local dataset saved at ${new Date(record.createdAt).toLocaleString()}.`)
    } catch {
      setStatus('error')
      setMessage('Unable to load data from local database.')
    }
  }

  const handleLoadDataset = (dataset) => {
    const normalizedLoadedRows = (dataset.rows || []).map((row) =>
      row.map((cell) => normalizeCellValue(cell)),
    )

    setHeaders(dataset.headers || [])
    setTableData(normalizedLoadedRows)
    setCellErrors(collectCellErrors(normalizedLoadedRows))
    setIsPreviewVisible(true)
    setStatus('success')
    setMessage(
      `Loaded ${dataset.fileName || 'excel-data'} saved at ${new Date(dataset.createdAt).toLocaleString()}.`,
    )
  }

  const handleDeleteDataset = async (dataset) => {
    const confirmed = window.confirm(
      `Delete "${dataset.fileName || 'excel-data'}" saved at ${new Date(dataset.createdAt).toLocaleString()}?`,
    )

    if (!confirmed) {
      return
    }

    try {
      await deleteDatasetFromDb(dataset.id)
      const allRecords = await loadAllDatasetsFromDb()
      setSavedDatasets(allRecords)
      setStatus('success')
      setMessage('Deleted dataset from local database.')
    } catch {
      setStatus('error')
      setMessage('Unable to delete dataset from local database.')
    }
  }

  const handleRefreshSavedDatasets = async () => {
    try {
      const allRecords = await loadAllDatasetsFromDb()
      setSavedDatasets(allRecords)
    } catch {
      setStatus('error')
      setMessage('Unable to fetch dataset list from local database.')
    }
  }

  useEffect(() => {
    handleRefreshSavedDatasets()
  }, [])

  return (
    <main className="excel-shell">
      <div className="bg-orb bg-orb-left" aria-hidden="true"></div>
      <div className="bg-orb bg-orb-right" aria-hidden="true"></div>

      <section className="excel-card" aria-labelledby="upload-title">
        <p className="chip">Agilysys</p>
        <h1 id="upload-title">SWS Item Tool</h1>
      
        <p className="subtitle">
          Select a file to continue. Accepted formats are .xls, .xlsx, and .csv.
        </p>

        <div className="upload-row">
          <label className="file-label" htmlFor="excel-file">
            Choose Excel File
          </label>
          <input
            id="excel-file"
            className="file-input"
            type="file"
            accept=".xls,.xlsx,.csv"
            onChange={handleFileChange}
          />
        </div>

        <div className="file-meta" role="status" aria-live="polite">
          <span className="meta-label">Selected file:</span>
          <span className="meta-value">{selectedFile ? selectedFile.name : 'None'}</span>
        </div>

        <button
          type="button"
          className="upload-btn"
          onClick={handleUpload}
        >
          Upload
        </button>

        <button
          type="button"
          className="secondary-btn"
          onClick={handleLoadFromLocalDb}
        >
          Load From Local DB
        </button>

        <button
          type="button"
          className="secondary-btn list"
          onClick={handleRefreshSavedDatasets}
        >
          Refresh Saved Datasets
        </button>

        <p className={`status ${status}`}>{message}</p>

        <div className="dataset-list-wrap" aria-live="polite">
          <h3>Saved Datasets</h3>
          <input
            type="text"
            className="dataset-search"
            placeholder="Search by file name"
            value={datasetSearchTerm}
            onChange={(event) => setDatasetSearchTerm(event.target.value)}
            aria-label="Search saved datasets by file name"
          />

          {savedDatasets.length === 0 ? (
            <p className="dataset-empty">No saved datasets yet.</p>
          ) : filteredDatasets.length === 0 ? (
            <p className="dataset-empty">No dataset matches your search.</p>
          ) : (
            <ul className="dataset-list">
              {filteredDatasets.map((dataset) => (
                <li key={dataset.id}>
                  <button
                    type="button"
                    className="dataset-load-btn"
                    onClick={() => handleLoadDataset(dataset)}
                    title="Load this saved dataset"
                  >
                    <span className="dataset-name">{dataset.fileName || 'excel-data'}</span>
                    <span className="dataset-time">
                      {new Date(dataset.createdAt).toLocaleString()}
                    </span>
                  </button>

                  <button
                    type="button"
                    className="dataset-delete-btn"
                    onClick={() => handleDeleteDataset(dataset)}
                    title="Delete this saved dataset"
                  >
                    Delete
                  </button>
                </li>
              ))}
            </ul>
          )}
        </div>
      </section>

      {isPreviewVisible && headers.length > 0 && (
        <section className="preview-card" aria-labelledby="preview-title">
          <div className="preview-head">
            <h2 id="preview-title">Excel Data Preview</h2>
            <div className="preview-meta">
              <span className="edit-hint">Click any cell to edit</span>
              <span className="row-count">Rows: {tableData.length}</span>
              <span className={`validation-count ${Object.keys(cellErrors).length ? 'error' : 'ok'}`}>
                Errors: {Object.keys(cellErrors).length}
              </span>
              <button type="button" className="save-btn add" onClick={handleAddNewLine}>
                Add
              </button>
              <button type="button" className="save-btn sort" onClick={handleSortColumnB}>
                Sort By Company
              </button>
              <button type="button" className="save-btn sort" onClick={handleSortByErrorRow}>
                Sort Error Rows
              </button>
              <button type="button" className="save-btn" onClick={handleConvertRow6ToUppercase}>
                Convert Row 6 to Uppercase
              </button>
              <button type="button" className="save-btn db" onClick={handleSaveToLocalDb}>
                Save To Local DB
              </button>
              <button type="button" className="save-btn" onClick={handleSaveCsv}>
                Save CSV
              </button>
              <button type="button" className="save-btn" onClick={handleSaveExcel}>
                Save Excel
              </button>
            </div>
          </div>

          <div className="table-wrap">
            <table className="excel-table">
              <thead>
                <tr>
                  <th className="row-index-header" aria-label="Row number">#</th>
                  {headers.map((header, index) => (
                    <th
                      key={`${header}-${index}`}
                      className={header === 'SKIP BPO CHECK FOR PR' ? 'header-highlight-green' : ''}
                    >
                      {header}
                    </th>
                  ))}
                  <th className="row-action-header" aria-label="Row actions">Action</th>
                </tr>
              </thead>
              <tbody>
                {tableData.map((row, rowIndex) => {
                  const rowHasError = row.some((_, cellIndex) =>
                    Boolean(cellErrors[`${rowIndex}-${cellIndex}`]),
                  )

                  return (
                  <tr key={`row-${rowIndex}`} className={rowHasError ? 'row-error' : ''}>
                    <td className="row-index-cell" aria-label={`Row ${rowIndex + 1}`}>
                      {rowIndex + 1}
                    </td>
                    {row.map((cell, cellIndex) => {
                      const errorMessage = cellErrors[`${rowIndex}-${cellIndex}`]
                      const cellText = String(cell ?? '')
                      const inputWidthCh = Math.max(10, Math.min(80, cellText.length + 2))
                      const isColumnA = cellIndex === 0
                      const isReadOnlyRow = rowIndex < 4
                      const isReadOnlyCell = isColumnA || isReadOnlyRow

                      return (
                        <td
                          key={`cell-${rowIndex}-${cellIndex}`}
                          className={errorMessage ? 'cell-error' : ''}
                          title={errorMessage || ''}
                        >
                          <input
                            className={`table-cell-input ${errorMessage ? 'input-error' : ''} ${isReadOnlyCell ? 'input-readonly' : ''}`}
                            type="text"
                            value={cellText}
                            style={{ width: `${inputWidthCh}ch` }}
                            onChange={
                              isReadOnlyCell
                                ? undefined
                                : (event) => handleCellChange(rowIndex, cellIndex, event.target.value)
                            }
                            readOnly={isReadOnlyCell}
                            aria-label={`${headers[cellIndex]} row ${rowIndex + 1}`}
                            aria-readonly={isReadOnlyCell}
                            aria-invalid={Boolean(errorMessage)}
                          />
                        </td>
                      )
                    })}

                    <td className="row-action-cell">
                      {rowIndex >= 4 ? (
                        <button
                          type="button"
                          className="row-delete-btn"
                          onClick={() => handleDeleteRow(rowIndex)}
                          aria-label={`Delete row ${rowIndex + 1}`}
                        >
                          Delete
                        </button>
                      ) : null}
                    </td>
                  </tr>
                )})}
              </tbody>
            </table>
          </div>
        </section>
      )}
    </main>
  )
}

export default App
