'use client'

import React, { useState, useRef, useEffect } from 'react'
import * as XLSX from 'xlsx'
import ExcelJS from 'exceljs'
import { parse, isValid, format } from 'date-fns'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Download, FileSpreadsheet, AlertCircle, CheckCircle2, Eye, Loader2, Check, ChevronLeft, ChevronRight, FileUp, Columns, Sparkles, FileText, RotateCcw, Coins, ShieldCheck } from 'lucide-react'
import ColumnMapper from './column-mapper'

interface DonationRecord {
  [key: string]: any
}

interface ColumnMapping {
  [key: string]: string | string[] | null
}

interface GiftAidConfig {
  column: string | null
  valuesToKeep: string[]
}

interface ProcessedRecord {
  original: any
  processed: any
  isValid: boolean
}

interface ProcessingAnalytics {
  titlesFilled: number
  namesFixed: number
  namesSplitFromEmail: number
  addressesShortened: number
  postcodesCorrected: number
  datesFormatted: number
  totalCellsModified: number
}

interface PreviewData {
  sheetNames: string[]
  totalRecords: number
  eligibleRecords: number
  beforeAfter: ProcessedRecord[]
  columns: string[]
  rawRecords: any[]
}

export default function DonationProcessor() {
  const [file, setFile] = useState<File | null>(null)
  const [preview, setPreview] = useState<PreviewData | null>(null)
  const [columnMapping, setColumnMapping] = useState<ColumnMapping | null>(null)
  const [giftAidConfig, setGiftAidConfig] = useState<GiftAidConfig | null>(null)
  const [showMapper, setShowMapper] = useState(false)
  const [processing, setProcessing] = useState(false)
  const [progressTarget, setProgressTarget] = useState(0)
  const [processingProgress, setProcessingProgress] = useState(0)
  const [loading, setLoading] = useState(false)
  const [previewPage, setPreviewPage] = useState(1)
  const [downloadUrls, setDownloadUrls] = useState<{
    cleanedData: string | null
    hmrcFiles: { name: string; url: string }[]
    outputFiles: { name: string; url: string }[]
  }>({ cleanedData: null, hmrcFiles: [], outputFiles: [] })
  const [result, setResult] = useState<{
    totalRecords: number
    eligibleRecords: number
    validRecords: number
    invalidRecords: number
    filteredRecords: number
    totalAmountReviewed: number
    totalGiftAidValue: number
    outputFiles: number
    analytics: ProcessingAnalytics
  } | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [recordsPerPage, setRecordsPerPage] = useState(10)
  const analyticsRef = useRef<HTMLDivElement>(null)
  const progressRef = useRef<number>(0)
  const animationRef = useRef<ReturnType<typeof requestAnimationFrame> | null>(null)

  // Smooth progress bar animation
  useEffect(() => {
    const animate = () => {
      const current = progressRef.current
      const diff = progressTarget - current
      if (Math.abs(diff) < 0.5) {
        progressRef.current = progressTarget
        setProcessingProgress(progressTarget)
        return
      }
      // Ease towards target: fast at first, slows down as it approaches
      const step = diff * 0.08
      progressRef.current = current + step
      setProcessingProgress(Math.round(progressRef.current))
      animationRef.current = requestAnimationFrame(animate)
    }

    if (progressTarget > progressRef.current || progressTarget === 0) {
      if (progressTarget === 0) {
        progressRef.current = 0
        setProcessingProgress(0)
        return
      }
      animationRef.current = requestAnimationFrame(animate)
    }

    return () => {
      if (animationRef.current) cancelAnimationFrame(animationRef.current)
    }
  }, [progressTarget])

  // Apply column mapping to records
  const applyColumnMapping = (record: any, mapping: ColumnMapping): DonationRecord => {
    const mapped: DonationRecord = {}

    Object.entries(mapping).forEach(([targetCol, sourceCol]) => {
      if (!sourceCol) return

      // Map to expected internal column names
      const internalMapping: { [key: string]: string } = {
        'Title': 'Title Prefix',
        'First Name': 'First Name',
        'Last Name': 'Last Name',
        'Email': 'Email',
        'Address': 'Gift Aid Address 1',
        'Postcode': 'Gift Aid Postal Code',
        'Date': 'Last Donation Date',
        'Amount': 'Total Donation Amount',
        'Gift Aid': 'Gift Aid'
      }

      const internalKey = internalMapping[targetCol] || targetCol

      // Handle multiple columns (e.g., Address)
      if (Array.isArray(sourceCol)) {
        const values = sourceCol.map(col => String(record[col] || '')).filter(v => v.trim())
        mapped[internalKey] = values.join(', ')
      } else if (record[sourceCol] !== undefined) {
        mapped[internalKey] = record[sourceCol]
      }
    })

    return mapped
  }

  const VALID_TITLES = ['Mr', 'Mrs', 'Miss', 'Ms', 'Dr', 'Prof']



  const assignTitle = (record: any): { title: string; cleanedRecord: any } => {
    let title = String(record['Title Prefix'] || '').trim()
    let firstName = String(record['First Name'] || '').trim()
    let lastName = String(record['Last Name'] || '').trim()

    const extractTitle = (str: string): { found: string | null; cleaned: string } => {
      if (!str) return { found: null, cleaned: str }

      const parts = str.split(/[\s,.]+/)
      for (const part of parts) {
        const match = VALID_TITLES.find((t: string) => t.toLowerCase() === part.toLowerCase())
        if (match) {
          const regex = new RegExp(`\\b${match}\\b\\.?`, 'gi')
          const cleaned = str.replace(regex, '').replace(/\s+/g, ' ').trim()
          return { found: match, cleaned }
        }
      }
      return { found: null, cleaned: str }
    }

    // If title is missing or invalid, try to extract from names
    const existingTitleMatch = VALID_TITLES.find((t: string) => t.toLowerCase() === title.toLowerCase())
    if (!title || !existingTitleMatch) {
      const fromFirst = extractTitle(firstName)
      if (fromFirst.found) {
        title = fromFirst.found
        firstName = fromFirst.cleaned
      } else {
        const fromLast = extractTitle(lastName)
        if (fromLast.found) {
          title = fromLast.found
          lastName = fromLast.cleaned
        }
      }
    } else {
      title = existingTitleMatch
    }

    if (!VALID_TITLES.includes(title)) {
      title = ''
    }

    const cleanedRecord = {
      ...record,
      'Title Prefix': title,
      'First Name': firstName,
      'Last Name': lastName
    }

    return { title, cleanedRecord }
  }

  const extractNameFromEmail = (email: string): { firstName: string; lastName: string } | null => {
    if (!email || !email.includes('@')) return null

    const localPart = email.split('@')[0]

    // Remove numbers but keep track of position for better parsing
    // e.g., rashi2zzz -> extract zzz as potential last name
    // rashi.zzz or rashi_zzz or rashi-zzz -> extract zzz as last name

    // First try to split by common delimiters
    const delimiters = /[._-]/
    if (delimiters.test(localPart)) {
      // Has delimiters like . _ or -
      const parts = localPart.split(delimiters).filter(p => p.length > 0)

      // Clean each part by removing numbers
      const cleanedParts = parts.map(p => p.replace(/\d+/g, '')).filter(p => p.length > 0)

      if (cleanedParts.length >= 2) {
        return {
          firstName: cleanedParts[0],
          lastName: cleanedParts[cleanedParts.length - 1] // Take last part as last name
        }
      } else if (cleanedParts.length === 1) {
        return { firstName: cleanedParts[0], lastName: cleanedParts[0][0] }
      }
    }

    // If no delimiters, try to extract pattern like "rashi2zzz"
    // Look for number followed by letters at the end
    const numberPattern = /^([a-zA-Z]+)\d+([a-zA-Z]+)$/
    const match = localPart.match(numberPattern)

    if (match && match[1] && match[2]) {
      // Found pattern like rashi2zzz
      return {
        firstName: match[1],
        lastName: match[2]
      }
    }

    // If still nothing, just remove all numbers and use as first name
    const cleaned = localPart.replace(/\d+/g, '')
    if (cleaned.length > 1) {
      return { firstName: cleaned, lastName: cleaned[0] }
    }

    return null
  }

  const isEligibleDonation = (record: any, config: GiftAidConfig | null): boolean => {
    if (!config || !config.column) {
      // Fallback to old logic if no config
      const giftAid = String(record['Gift Aid'] || record['DonationTaxEffective'] || '').toLowerCase().trim()
      const taxEligible = String(record['Tax Eligible'] || '').toLowerCase().trim()

      return giftAid.includes('tax effective') ||
        giftAid === 'yes' || giftAid === 'y' || giftAid === 'true' ||
        taxEligible === 'yes' || taxEligible === 'y' || taxEligible === 'true'
    }

    // Use configured Gift Aid column and values
    const value = String(record[config.column] || '').trim()
    return config.valuesToKeep.includes(value)
  }

  const needsColoringPostcode = (pc: any): boolean => {
    if (!pc) return true
    const pcStr = String(pc).trim()
    if (pcStr === '') return true
    // Flag postcodes that are too short (< 4) or too long (> 8)
    if (pcStr.length < 4 || pcStr.length > 8) return true
    if (pcStr.substring(0, 5).match(/^\d+$/)) return true
    return false
  }

  const needsColoringAddress = (addr: any): boolean => {
    if (!addr) return true
    const addrStr = String(addr).trim()
    if (addrStr === '' || addrStr.match(/^\d+$/)) return true
    return false
  }

  const formatPostcode = (pc: any): string | null => {
    if (!pc) return null
    let pcStr = String(pc).trim().toUpperCase()
    if (pcStr === '') return null

    pcStr = pcStr.replace(/\s+/g, '')

    if (pcStr.length >= 5) {
      const inward = pcStr.slice(-3)
      const outward = pcStr.slice(0, -3)
      return `${outward} ${inward}`
    }

    return pcStr
  }

  const formatDate = (date: any): string | null => {
    if (!date) return null

    try {
      let d: Date | null = null

      // Handle Excel serial date numbers
      if (typeof date === 'number') {
        const excelEpoch = new Date(1899, 11, 30)
        d = new Date(excelEpoch.getTime() + date * 86400000)
      }
      // Handle Date objects
      else if (date instanceof Date) {
        d = date
      }
      // Handle string dates - try multiple formats
      else if (typeof date === 'string') {
        const dateStr = date.trim()

        // Try common date formats
        const formats = [
          'MM/dd/yyyy',  // 12/15/2025
          'M/d/yyyy',    // 12/5/2025
          'dd/MM/yyyy',  // 15/12/2025
          'd/M/yyyy',    // 5/12/2025
          'yyyy-MM-dd',  // 2025-12-15
          'dd-MM-yyyy',  // 15-12-2025
          'MM-dd-yyyy',  // 12-15-2025
        ]

        for (const fmt of formats) {
          const parsed = parse(dateStr, fmt, new Date())
          if (isValid(parsed)) {
            d = parsed
            break
          }
        }

        // If none of the formats worked, try native Date parsing
        if (!d) {
          d = new Date(dateStr)
          if (isNaN(d.getTime())) {
            d = null
          }
        }
      }

      // If we have a valid date, format it
      if (d && !isNaN(d.getTime())) {
        return format(d, 'dd/MM/yyyy')
      }

      return null
    } catch {
      return null
    }
  }

  const updateNames = (record: DonationRecord): DonationRecord => {
    let firstName = String(record['First Name'] || '').trim()
    let lastName = String(record['Last Name'] || '').trim()
    const email = String(record['Email'] || '').trim()

    // Both names missing - try email first
    if (!firstName && !lastName) {
      if (email) {
        const emailNames = extractNameFromEmail(email)
        if (emailNames) {
          return {
            ...record,
            'First Name': emailNames.firstName,
            'Last Name': emailNames.lastName
          }
        }
      }
      return record
    }

    // Only first name exists - try to extract last name from email
    if (firstName && !lastName) {
      const parts = firstName.split(/\s+/)
      if (parts.length > 1) {
        return {
          ...record,
          'First Name': parts[0],
          'Last Name': parts.slice(1).join(' ')
        }
      } else if (email) {
        const emailNames = extractNameFromEmail(email)
        if (emailNames && emailNames.lastName) {
          return { ...record, 'Last Name': emailNames.lastName }
        }
      }
    }

    // Only last name exists - try to extract first name from email
    if (lastName && !firstName) {
      const parts = lastName.split(/\s+/)
      if (parts.length > 1) {
        return {
          ...record,
          'First Name': parts[0],
          'Last Name': parts.slice(1).join(' ')
        }
      } else if (email) {
        const emailNames = extractNameFromEmail(email)
        if (emailNames && emailNames.firstName) {
          return { ...record, 'First Name': emailNames.firstName }
        }
      }
    }

    return record
  }

  const processRecord = (record: any, config: GiftAidConfig | null, analytics?: ProcessingAnalytics): ProcessedRecord & { changesApplied?: string[] } => {
    if (!isEligibleDonation(record, config)) {
      return {
        original: record,
        processed: null,
        isValid: false,
        changesApplied: []
      }
    }

    const originalRecord = { ...record }
    const { title, cleanedRecord: withTitle } = assignTitle(record)
    const updated = updateNames(withTitle)

    // Track changes made to this specific record
    const changesApplied: string[] = []

    // Track analytics and changes
    if (analytics) {
      // Track title changes - only if we actually found/fixed a title
      const originalTitle = String(originalRecord['Title Prefix'] || '').trim()
      if (title && originalTitle !== title) {
        analytics.titlesFilled++
        analytics.totalCellsModified++
        changesApplied.push('Title')
      }

      // Track name fixes
      const originalFirstName = String(originalRecord['First Name'] || '').trim()
      const originalLastName = String(originalRecord['Last Name'] || '').trim()
      const newFirstName = String(updated['First Name'] || '').trim()
      const newLastName = String(updated['Last Name'] || '').trim()
      const email = String(originalRecord['Email'] || '').trim()

      if (!originalFirstName && !originalLastName && (newFirstName === 'A' && newLastName === 'Anonymous')) {
        analytics.namesFixed++
        analytics.totalCellsModified += 2
        changesApplied.push('Names (Anonymous)')
      } else if (originalFirstName !== newFirstName || originalLastName !== newLastName) {
        // Check if the change actually came from email
        let actuallyFromEmail = false
        if (email) {
          const emailNames = extractNameFromEmail(email)
          if (emailNames) {
            // Check if the new names match what we extracted from email
            if ((!originalFirstName && newFirstName === emailNames.firstName) ||
              (!originalLastName && newLastName === emailNames.lastName)) {
              actuallyFromEmail = true
              analytics.namesSplitFromEmail++
            }
          }
        }

        analytics.namesFixed++
        if (originalFirstName !== newFirstName) {
          analytics.totalCellsModified++
          changesApplied.push('First Name')
        }
        if (originalLastName !== newLastName) {
          analytics.totalCellsModified++
          changesApplied.push('Last Name')
        }
      }

      // Track postcode corrections
      const originalPostcode = String(originalRecord['Gift Aid Postal Code'] || '').trim()
      const formattedPostcode = formatPostcode(originalRecord['Gift Aid Postal Code'])
      if (originalPostcode && originalPostcode !== formattedPostcode) {
        analytics.postcodesCorrected++
        analytics.totalCellsModified++
        changesApplied.push('Postcode')
      }

      // Track date formatting - only if the date was actually reformatted
      const originalDate = originalRecord['Last Donation Date']
      const formattedDate = formatDate(originalDate)
      const originalDateStr = String(originalDate || '').trim()

      // Check if date was actually changed (not already in dd/MM/yyyy format)
      if (originalDate && formattedDate && originalDateStr !== formattedDate) {
        // Date was reformatted
        analytics.datesFormatted++
        analytics.totalCellsModified++
        changesApplied.push('Date')
      }

      // Track address shortening (if it was too long)
      const address = String(originalRecord['Gift Aid Address 1'] || '').trim()
      if (address && address.length > 50) {
        analytics.addressesShortened++
        analytics.totalCellsModified++
        changesApplied.push('Address')
      }
    }

    const finalFirstName = String(updated['First Name'] || '').trim()
    const finalLastName = String(updated['Last Name'] || '').trim()

    const needsReview =
      needsColoringPostcode(record['Gift Aid Postal Code']) ||
      needsColoringAddress(record['Gift Aid Address 1']) ||
      !record['Last Donation Date'] ||
      !finalFirstName ||
      !finalLastName

    const processed = {
      'Title': title,
      'First Name': String(updated['First Name'] || ''),
      'Last Name': String(updated['Last Name'] || ''),
      'Address': String(updated['Gift Aid Address 1'] || ''),
      'Postcode': needsReview ? String(record['Gift Aid Postal Code'] || '') : formatPostcode(record['Gift Aid Postal Code']),
      'Donation Date': formatDate(record['Last Donation Date']),
      'Donation Amount': Number(record['Total Donation Amount']) || 0,
      'needsReview': needsReview
    }

    return {
      original: record,
      processed: processed,
      isValid: !needsReview,
      changesApplied: changesApplied
    }
  }

  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      const selectedFile = e.target.files[0]
      setFile(selectedFile)
      setError(null)
      setResult(null)
      setColumnMapping(null)
      setShowMapper(false)
      setLoading(true)

      try {
        const data = await selectedFile.arrayBuffer()
        const workbook = XLSX.read(data, { type: 'array', sheetRows: 0 })

        let allRecords: any[] = []
        const sheetNames = workbook.SheetNames
        let allColumns: string[] = []

        workbook.SheetNames.forEach(sheetName => {
          const worksheet = workbook.Sheets[sheetName]

          // Get all columns including empty ones
          const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1')
          const cols: string[] = []

          // Read header row to get all column names
          for (let C = range.s.c; C <= range.e.c; ++C) {
            const cellAddress = XLSX.utils.encode_cell({ r: range.s.r, c: C })
            const cell = worksheet[cellAddress]
            cols.push(cell ? String(cell.v) : `Column${C}`)
          }

          // Merge columns from all sheets
          cols.forEach(col => {
            if (!allColumns.includes(col)) {
              allColumns.push(col)
            }
          })

          // Convert to JSON with all columns preserved
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '' })
          allRecords = allRecords.concat(jsonData)
        })

        const columns = allColumns.length > 0 ? allColumns : (allRecords.length > 0 ? Object.keys(allRecords[0]) : [])

        // Check if columns match expected format
        const hasStandardColumns = ['First Name', 'Last Name', 'Gift Aid'].every(col =>
          columns.some(c => c.toLowerCase().includes(col.toLowerCase().replace(' ', '')))
        )

        if (!hasStandardColumns) {
          // Show mapper
          setShowMapper(true)
        }

        setPreview({
          sheetNames,
          totalRecords: allRecords.length,
          eligibleRecords: 0,
          beforeAfter: [],
          columns,
          rawRecords: allRecords
        })
      } catch (err) {
        setError('Error reading file: ' + (err as Error).message)
      } finally {
        setLoading(false)
      }
    }
  }

  const handleMappingComplete = (mapping: ColumnMapping, config: GiftAidConfig) => {
    setColumnMapping(mapping)

    // Update config to use internal Gift Aid column name after mapping
    const updatedConfig: GiftAidConfig = {
      column: 'Gift Aid', // After mapping, always use internal name
      valuesToKeep: config.valuesToKeep
    }
    setGiftAidConfig(updatedConfig)
    setShowMapper(false)
    setPreviewPage(1) // Reset to first page

    // Re-process ALL records for preview with mapping
    if (preview) {
      const mappedRecords = preview.rawRecords.map(rec => applyColumnMapping(rec, mapping))
      const allProcessed = mappedRecords.map(rec => processRecord(rec, updatedConfig))
      const eligibleCount = mappedRecords.filter(rec => isEligibleDonation(rec, updatedConfig)).length

      setPreview({
        ...preview,
        eligibleRecords: eligibleCount,
        beforeAfter: allProcessed
      })
    }
  }

  const handleSkipMapper = () => {
    setShowMapper(false)

    if (preview) {
      const processedSamples = preview.rawRecords.slice(0, 10).map(rec => processRecord(rec, null))
      const eligibleCount = preview.rawRecords.filter(rec => isEligibleDonation(rec, null)).length

      setPreview({
        ...preview,
        eligibleRecords: eligibleCount,
        beforeAfter: processedSamples
      })
    }
  }

  const processFile = async () => {
    if (!file || !preview) {
      setError('Please select a file')
      return
    }

    setProcessing(true)
    setError(null)

    try {
      let allRecords = preview.rawRecords

      // Apply column mapping if exists
      if (columnMapping) {
        allRecords = allRecords.map(rec => applyColumnMapping(rec, columnMapping))
      }

      // Initialize analytics
      const analytics: ProcessingAnalytics = {
        titlesFilled: 0,
        namesFixed: 0,
        namesSplitFromEmail: 0,
        addressesShortened: 0,
        postcodesCorrected: 0,
        datesFormatted: 0,
        totalCellsModified: 0
      }

      setProgressTarget(15)
      await new Promise(r => setTimeout(r, 100))

      const eligibleRecords = allRecords.filter(rec => isEligibleDonation(rec, giftAidConfig))
      const ineligibleRecords = allRecords.filter(rec => !isEligibleDonation(rec, giftAidConfig))
      setProgressTarget(30)
      await new Promise(r => setTimeout(r, 100))

      const processedRecords = eligibleRecords.map(rec => processRecord(rec, giftAidConfig, analytics))
      setProgressTarget(45)

      // Only valid records (those that don't need review) - for "Correct Data" sheet and HMRC template
      const allProcessedForCorrect = processedRecords
        .filter(r => r.isValid && r.processed !== null)
        .map(r => r.processed)

      // Only records that need review - for "Needs Review" sheet
      const recordsNeedingReview = processedRecords
        .filter(r => !r.isValid && r.processed !== null)
        .map(r => r.processed)

      const totalAmountReviewed = eligibleRecords.reduce((sum, rec) => sum + (Number(rec['Total Donation Amount']) || 0), 0)
      const totalGiftAidValue = allProcessedForCorrect.reduce((sum, rec) => sum + (Number((rec as any)['Donation Amount']) || 0), 0)

      const cleanedData = allProcessedForCorrect.map(r => {
        const { needsReview, ...rest } = r as any
        return rest
      })

      // Find earliest date
      let earliestDate: Date | null = null;
      cleanedData.forEach(record => {
        const dateStr = (record as any)['Donation Date'];
        if (dateStr) {
          const d = parse(dateStr, 'dd/MM/yyyy', new Date());
          if (isValid(d)) {
            if (!earliestDate || d < earliestDate) {
              earliestDate = d;
            }
          }
        }
      });
      const earliestDateStr = earliestDate ? format(earliestDate, 'dd/MM/yyyy') : '';

      const chunks: any[][] = []
      for (let i = 0; i < cleanedData.length; i += 1000) {
        chunks.push(cleanedData.slice(i, i + 1000))
      }

      setProgressTarget(52)
      await new Promise(r => setTimeout(r, 100))

      // Load the HMRC template from public folder
      const templateResponse = await fetch('/tempxl.xlsx')
      const templateArrayBuffer = await templateResponse.arrayBuffer()
      const templateWorkbook = XLSX.read(templateArrayBuffer, { type: 'array' })

      setProgressTarget(58)

      // Create workbook with sheets
      const resultWorkbook = XLSX.utils.book_new()

      // "Correct Data" sheet - contains ALL processed records (valid + needs review)
      const cleanedSheet = XLSX.utils.json_to_sheet(cleanedData)
      XLSX.utils.book_append_sheet(resultWorkbook, cleanedSheet, 'Correct Data')

      // "Needs Review" sheet - only records that need manual verification
      if (recordsNeedingReview.length > 0) {
        const reviewSheet = XLSX.utils.json_to_sheet(recordsNeedingReview.map(r => {
          const { needsReview, ...rest } = r
          return rest
        }))
        XLSX.utils.book_append_sheet(resultWorkbook, reviewSheet, 'Needs Review')
      }

      // "Ineligible Records" sheet - records filtered out by logic
      if (ineligibleRecords.length > 0) {
        const filteredSheet = XLSX.utils.json_to_sheet(ineligibleRecords)
        XLSX.utils.book_append_sheet(resultWorkbook, filteredSheet, 'Ineligible Records')
      }

      // Create a detailed sheet with auto-fixed records highlighted
      const allProcessedData = processedRecords.map((r, idx) => {
        const record = r.processed || {}
        const changes = r.changesApplied || []

        return {
          'Row #': idx + 1,
          'Status': changes.length > 0 ? 'Auto-Fixed' : 'OK',
          'Changes Applied': changes.length > 0 ? changes.join(', ') : 'None',
          'Title': record.Title || '',
          'First Name': record['First Name'] || '',
          'Last Name': record['Last Name'] || '',
          'Address': record.Address || '',
          'Postcode': record.Postcode || '',
          'Donation Date': record['Donation Date'] || '',
          'Donation Amount': record['Donation Amount'] || 0
        }
      })

      const fixedSheet = XLSX.utils.json_to_sheet(allProcessedData)

      // Apply cell styling for rows that were auto-fixed
      const range = XLSX.utils.decode_range(fixedSheet['!ref'] || 'A1')

      for (let R = range.s.r + 1; R <= range.e.r; ++R) {
        const statusCell = fixedSheet[XLSX.utils.encode_cell({ r: R, c: 1 })]

        // Color rows that were auto-fixed (Status column contains "Auto-Fixed")
        if (statusCell && statusCell.v && String(statusCell.v).includes('Auto-Fixed')) {
          for (let C = range.s.c; C <= range.e.c; ++C) {
            const cellAddress = XLSX.utils.encode_cell({ r: R, c: C })
            if (!fixedSheet[cellAddress]) continue

            // Add cell styling (light green background for auto-fixed rows)
            if (!fixedSheet[cellAddress].s) fixedSheet[cellAddress].s = {}
            fixedSheet[cellAddress].s = {
              fill: { fgColor: { rgb: "D4EDDA" } }, // Light green
              font: { color: { rgb: "155724" }, bold: true } // Dark green text
            }
          }
        }
      }

      // Set column widths
      fixedSheet['!cols'] = [
        { wch: 8 },  // Row #
        { wch: 12 }, // Status
        { wch: 30 }, // Changes Applied
        { wch: 8 },  // Title
        { wch: 15 }, // First Name
        { wch: 15 }, // Last Name
        { wch: 30 }, // Address
        { wch: 12 }, // Postcode
        { wch: 12 }, // Date
        { wch: 12 }  // Amount
      ]

      XLSX.utils.book_append_sheet(resultWorkbook, fixedSheet, 'Auto-Fixed Records')

      // Convert cleaned_data workbook to blob and create download URL
      const wbout = XLSX.write(resultWorkbook, { bookType: 'xlsx', type: 'array' })
      const cleanedDataBlob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
      const cleanedDataUrl = window.URL.createObjectURL(cleanedDataBlob)

      setProgressTarget(65)
      await new Promise(r => setTimeout(r, 100))

      // Load the HMRC template and populate it with data using ExcelJS for perfect formatting preservation
      const hmrcUrls: { name: string; url: string }[] = []

      for (let chunkIndex = 0; chunkIndex < chunks.length; chunkIndex++) {
        const chunk = chunks[chunkIndex]

        // Load template with ExcelJS
        const templateResponse = await fetch('/tempxl.xlsx')
        const templateBuffer = await templateResponse.arrayBuffer()

        const workbook = new ExcelJS.Workbook()
        await workbook.xlsx.load(templateBuffer)

        const worksheet = workbook.getWorksheet(1) // Get first worksheet

        if (!worksheet) {
          throw new Error('Template worksheet not found')
        }

        // Fill earliest date in D13 (Row 13, Column 4)
        if (earliestDateStr) {
          worksheet.getCell('D13').value = earliestDateStr;
        }

        // Insert data starting from row 25, column C
        const startRow = 25
        const startCol = 3 // Column C (1-indexed in ExcelJS)

        chunk.forEach((record, idx) => {
          const rowNumber = startRow + idx
          const row = worksheet.getRow(rowNumber)

          // Set values in columns C through K
          row.getCell(startCol).value = record.Title || ''                      // C: Title
          row.getCell(startCol + 1).value = record['First Name'] || ''          // D: First name
          row.getCell(startCol + 2).value = record['Last Name'] || ''           // E: Last name
          row.getCell(startCol + 3).value = record.Address || ''                // F: House name or number
          row.getCell(startCol + 4).value = record.Postcode || ''               // G: Postcode
          row.getCell(startCol + 5).value = ''                                  // H: Aggregated donations
          row.getCell(startCol + 6).value = ''                                  // I: Sponsored event
          row.getCell(startCol + 7).value = record['Donation Date'] || ''       // J: Donation date
          row.getCell(startCol + 8).value = record['Donation Amount'] || 0      // K: Amount

          row.commit()
        })

        // Generate buffer and create blob URL
        const buffer = await workbook.xlsx.writeBuffer()
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
        const url = window.URL.createObjectURL(blob)

        hmrcUrls.push({
          name: `HMRC_submission_${chunkIndex + 1}.xlsx`,
          url: url
        })

        setProgressTarget(65 + (chunkIndex + 1) * (25 / chunks.length))
        await new Promise(r => setTimeout(r, 100))
      }

      // Also generate simple output sheets
      const outputUrls: { name: string; url: string }[] = []

      chunks.forEach((chunk, index) => {
        const chunkWorkbook = XLSX.utils.book_new()
        const chunkSheet = XLSX.utils.json_to_sheet(chunk)
        XLSX.utils.book_append_sheet(chunkWorkbook, chunkSheet, 'Sheet1')

        const wbout = XLSX.write(chunkWorkbook, { bookType: 'xlsx', type: 'array' })
        const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
        const url = window.URL.createObjectURL(blob)

        outputUrls.push({
          name: `output_sheet_${index + 1}.xlsx`,
          url: url
        })
      })

      setProgressTarget(100)
      await new Promise(r => setTimeout(r, 400))

      setDownloadUrls({
        cleanedData: cleanedDataUrl,
        hmrcFiles: hmrcUrls,
        outputFiles: outputUrls
      })

      setResult({
        totalRecords: allRecords.length,
        eligibleRecords: eligibleRecords.length,
        validRecords: allProcessedForCorrect.length,
        invalidRecords: recordsNeedingReview.length,
        filteredRecords: ineligibleRecords.length,
        totalAmountReviewed,
        totalGiftAidValue,
        outputFiles: chunks.length,
        analytics: analytics
      })

    } catch (err) {
      setError('Error processing file: ' + (err as Error).message)
    } finally {
      setProcessing(false)
    }
  }

  const handleReset = () => {
    // Clean up blob URLs
    if (downloadUrls.cleanedData) {
      window.URL.revokeObjectURL(downloadUrls.cleanedData)
    }
    downloadUrls.hmrcFiles.forEach(file => window.URL.revokeObjectURL(file.url))
    downloadUrls.outputFiles.forEach(file => window.URL.revokeObjectURL(file.url))

    setFile(null)
    setPreview(null)
    setColumnMapping(null)
    setGiftAidConfig(null)
    setShowMapper(false)
    setResult(null)
    setError(null)
    setPreviewPage(1)
    setDownloadUrls({ cleanedData: null, hmrcFiles: [], outputFiles: [] })
    setProgressTarget(0)
  }

  const exportAnalyticsToPDF = () => {
    if (!result || !analyticsRef.current) return

    // Create a new window for printing
    const printWindow = window.open('', '_blank')
    if (!printWindow) return

    const currentDate = new Date().toLocaleDateString('en-GB')

    printWindow.document.write(`
      <!DOCTYPE html>
      <html>
      <head>
        <title>Gift Aid Summary Report</title>
        <style>
          body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            display: flex;
            justify-content: center;
            padding: 40px 20px;
            background-color: #f8fafc;
            color: #475569;
          }
          .summary-card {
            background: white;
            width: 100%;
            max-width: 500px;
            padding: 32px;
            border-radius: 16px;
            border: 1px solid #e2e8f0;
            box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1);
          }
          .header {
            color: #1e3a8a;
            font-size: 20px;
            font-weight: 700;
            margin-bottom: 24px;
            margin-top: 0;
          }
          .row {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 12px;
            font-size: 15px;
            line-height: 1.5;
          }
          .label {
            color: #64748b;
          }
          .value {
            font-weight: 700;
            color: #0f172a;
          }
          .value.green {
            color: #10b981;
          }
          .value.red {
            color: #ef4444;
          }
          .divider {
            height: 1px;
            background-color: #f1f5f9;
            margin: 20px 0;
          }
          .section-title {
            font-size: 16px;
            font-weight: 700;
            color: #0f172a;
            margin: 24px 0 16px 0;
          }
          .warning-icon {
            color: #ef4444;
            margin-right: 8px;
            display: inline-flex;
            vertical-align: middle;
          }
          @media print {
            body { background: white; padding: 0; }
            .summary-card { border: 1px solid #d1fae5; box-shadow: none; border-radius: 16px; }
            .no-print { display: none; }
          }
        </style>
      </head>
      <body>
        <div class="summary-card">
          <h1 class="header">Boost Aid - Gift Aid Summary</h1>
          
          <div class="row">
            <span class="label">Total Donations Reviewed:</span>
            <span class="value">£${result.totalAmountReviewed.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 0 })}</span>
          </div>
          
          <div class="row">
            <span class="label">Data Lines:</span>
            <span class="value">${result.totalRecords.toLocaleString()}</span>
          </div>
          
          <div class="row">
            <span class="label">Errors Fixed:</span>
            <span class="value">${result.analytics.totalCellsModified.toLocaleString()}</span>
          </div>
          
          <div class="row">
            <span class="label">Valid Records:</span>
            <span class="value">${result.validRecords.toLocaleString()}</span>
          </div>
          
          <div class="row">
            <span class="label">Total Gift Aid Value:</span>
            <span class="value">£${result.totalGiftAidValue.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 0 })}</span>
          </div>
          
          <div class="row">
            <span class="label">Estimated Gift Aid Reclaimable:</span>
            <span class="value green">£${(result.totalGiftAidValue * 0.25).toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 0 })}</span>
          </div>

          <div class="divider"></div>

          <h2 class="section-title">Data Quality Check</h2>

          <div class="row">
            <span class="label">Records needing attention:</span>
            <span class="value red">${result.invalidRecords.toLocaleString()}</span>
          </div>

          <div class="row">
            <span class="label">Compliance Rate:</span>
            <span class="value green">
              ${result.eligibleRecords > 0 ? Math.round((result.validRecords / result.eligibleRecords) * 100) : 100}%
            </span>
          </div>
        </div>

        <div class="no-print" style="position: fixed; bottom: 20px; left: 0; right: 0; text-align: center;">
          <button onclick="window.print()" style="background: #1e3a8a; color: white; border: none; padding: 12px 24px; border-radius: 8px; font-size: 14px; cursor: pointer; font-weight: 600; box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1);">
            Print / Save as PDF
          </button>
        </div>
      </body>
      </html>
    `)

    printWindow.document.close()
  }

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-900 via-slate-800 to-slate-900">
      {/* Header */}
      <div className="bg-slate-800 border-b border-slate-700/50 text-white">
        <div className="max-w-7xl mx-auto px-6 py-6">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-4">
              <div className="w-11 h-11 bg-gradient-to-br from-violet-500 to-purple-600 rounded-xl flex items-center justify-center shadow-lg shadow-purple-500/20">
                <FileSpreadsheet className="w-5 h-5" />
              </div>
              <div>
                <h1 className="text-xl font-semibold tracking-tight">BoostAid</h1>
                <p className="text-slate-400 text-sm">HMRC Gift Aid Submission Tool</p>
              </div>
            </div>
            <Button
              onClick={async () => {
                await fetch('/api/auth/logout', { method: 'POST' })
                window.location.href = '/login'
              }}
              variant="outline"
              size="sm"
              className="border-slate-600 bg-slate-700/50 text-slate-300 hover:bg-slate-600 hover:text-white"
            >
              <svg className="w-4 h-4 mr-2" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" d="M15.75 9V5.25A2.25 2.25 0 0013.5 3h-6a2.25 2.25 0 00-2.25 2.25v13.5A2.25 2.25 0 007.5 21h6a2.25 2.25 0 002.25-2.25V15m3 0l3-3m0 0l-3-3m3 3H9" />
              </svg>
              Logout
            </Button>
          </div>
        </div>
      </div>

      <div className="max-w-7xl mx-auto px-6 py-6 space-y-6">
        {/* Upload Section */}
        {!preview && (
          <>
            <div className="bg-slate-800/50 backdrop-blur-sm rounded-2xl shadow-xl border border-slate-700/50 overflow-hidden">
              <div className="p-6">
                <div className="flex items-center gap-3 mb-6">
                  <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-violet-500 to-purple-600 text-white flex items-center justify-center font-bold shadow-lg shadow-purple-500/30">
                    1
                  </div>
                  <div>
                    <h2 className="text-lg font-semibold text-white">Upload Your Excel File</h2>
                    <p className="text-sm text-slate-400">Select a .xlsx or .xls file containing donation records</p>
                  </div>
                </div>

                <label className="block">
                  <div className="border-2 border-dashed border-slate-600 rounded-xl p-8 text-center hover:border-purple-400 hover:bg-purple-500/5 transition-all cursor-pointer group">
                    <div className="w-16 h-16 mx-auto mb-4 rounded-full bg-slate-700/50 flex items-center justify-center group-hover:bg-purple-500/20 transition-colors">
                      <FileUp className="w-8 h-8 text-slate-400 group-hover:text-purple-400 transition-colors" />
                    </div>
                    <div className="text-slate-300 font-medium mb-1">Drop your file here or click to browse</div>
                    <div className="text-sm text-slate-500">Supports Excel files (.xlsx, .xls)</div>
                    <Input
                      type="file"
                      accept=".xlsx,.xls"
                      onChange={handleFileChange}
                      className="hidden"
                      disabled={loading}
                    />
                  </div>
                </label>

                {loading && (
                  <div className="flex items-center justify-center gap-3 mt-6 py-4 bg-purple-500/10 rounded-xl border border-purple-500/20">
                    <Loader2 className="w-5 h-5 animate-spin text-purple-400" />
                    <span className="text-purple-300 font-medium">Reading your file...</span>
                  </div>
                )}
              </div>
            </div>

            {/* How It Works */}
            <div className="bg-slate-800/50 backdrop-blur-sm rounded-2xl shadow-xl border border-slate-700/50 p-6">
              <h3 className="text-sm font-semibold text-white mb-4 flex items-center gap-2">
                <Sparkles className="w-4 h-4 text-purple-400" />
                How It Works
              </h3>
              <div className="grid md:grid-cols-3 gap-6">
                <div className="flex gap-4">
                  <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-violet-500/20 to-purple-500/20 border border-purple-500/30 text-purple-400 flex items-center justify-center flex-shrink-0">
                    <Columns className="w-5 h-5" />
                  </div>
                  <div>
                    <div className="font-semibold text-white mb-1">1. Map Columns</div>
                    <p className="text-sm text-slate-400">Drag and drop to match your columns to required fields</p>
                  </div>
                </div>
                <div className="flex gap-4">
                  <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-violet-500/20 to-purple-500/20 border border-purple-500/30 text-purple-400 flex items-center justify-center flex-shrink-0">
                    <Eye className="w-5 h-5" />
                  </div>
                  <div>
                    <div className="font-semibold text-white mb-1">2. Preview Changes</div>
                    <p className="text-sm text-slate-400">Review transformations before processing</p>
                  </div>
                </div>
                <div className="flex gap-4">
                  <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-violet-500/20 to-purple-500/20 border border-purple-500/30 text-purple-400 flex items-center justify-center flex-shrink-0">
                    <Download className="w-5 h-5" />
                  </div>
                  <div>
                    <div className="font-semibold text-white mb-1">3. Download Files</div>
                    <p className="text-sm text-slate-400">Get cleaned data ready for HMRC submission</p>
                  </div>
                </div>
              </div>
            </div>
          </>
        )}

        {showMapper && preview && (
          <ColumnMapper
            detectedColumns={preview.columns}
            sampleData={preview.rawRecords}
            onMappingComplete={handleMappingComplete}
          />
        )}

        {preview && !showMapper && preview.beforeAfter.length > 0 && !result && (
          <>
            {/* Step 2 Header */}
            <div className="bg-slate-800/50 backdrop-blur-sm rounded-2xl shadow-xl border border-slate-700/50 p-5">
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-4">
                  <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-violet-500 to-purple-600 text-white flex items-center justify-center font-bold shadow-lg shadow-purple-500/30">
                    2
                  </div>
                  <div>
                    <h2 className="text-lg font-semibold text-white">Preview & Process</h2>
                    <p className="text-sm text-slate-400">{preview.eligibleRecords.toLocaleString()} eligible records ready</p>
                  </div>
                </div>
                <Button
                  onClick={processFile}
                  disabled={processing}
                  size="lg"
                  className="bg-gradient-to-r from-violet-500 to-purple-600 hover:from-violet-600 hover:to-purple-700 text-white shadow-lg shadow-purple-500/30 border-0"
                >
                  {processing ? (
                    <>
                      <Loader2 className="w-4 h-4 mr-2 animate-spin" />
                      Processing...
                    </>
                  ) : (
                    <>
                      <Check className="w-4 h-4 mr-2" />
                      Process All Records
                    </>
                  )}
                </Button>
              </div>

              {/* Progress Bar */}
              {processing && (
                <div className="mt-4">
                  <div className="text-sm text-slate-300 mb-2">Processing your data...</div>
                  <div className="w-full h-3 bg-slate-700 rounded-full overflow-hidden">
                    <div
                      className="h-full bg-gradient-to-r from-violet-500 to-purple-600 transition-all duration-500 ease-out"
                      style={{ width: `${processingProgress}% ` }}
                    />
                  </div>
                </div>
              )}
            </div>

            {/* Stats */}
            <div className="grid grid-cols-4 gap-4">
              <div className="bg-slate-800/50 backdrop-blur-sm rounded-xl shadow-lg border border-slate-700/50 p-4">
                <div className="text-xs font-medium text-slate-400 uppercase tracking-wide">Total</div>
                <div className="text-2xl font-bold text-white mt-1">{preview.totalRecords.toLocaleString()}</div>
              </div>
              <div className="bg-gradient-to-br from-violet-500/10 to-purple-500/10 backdrop-blur-sm rounded-xl border border-purple-500/30 p-4">
                <div className="text-xs font-medium text-purple-300 uppercase tracking-wide">Eligible</div>
                <div className="text-2xl font-bold text-purple-200 mt-1">{preview.eligibleRecords.toLocaleString()}</div>
              </div>
              <div className="bg-slate-800/50 backdrop-blur-sm rounded-xl shadow-lg border border-slate-700/50 p-4">
                <div className="text-xs font-medium text-slate-400 uppercase tracking-wide">Sheets</div>
                <div className="text-2xl font-bold text-white mt-1">{preview.sheetNames.length}</div>
              </div>
              <div className="bg-slate-700/30 backdrop-blur-sm rounded-xl border border-slate-600/50 p-4">
                <div className="text-xs font-medium text-slate-400 uppercase tracking-wide">Filtered</div>
                <div className="text-2xl font-bold text-slate-300 mt-1">{(preview.totalRecords - preview.eligibleRecords).toLocaleString()}</div>
              </div>
            </div>

            {/* Preview Table - Transformed Data Only */}
            <div className="bg-slate-800/50 backdrop-blur-sm rounded-2xl shadow-xl border border-slate-700/50 overflow-hidden">
              <div className="px-5 py-4 bg-slate-800/80 border-b border-slate-700/50 flex items-center justify-between">
                <div>
                  <h3 className="font-semibold text-white">Output Preview</h3>
                  <p className="text-xs text-slate-400 mt-0.5">How your data will look after processing</p>
                </div>
                {(() => {
                  const eligibleRecords = preview.beforeAfter.filter(i => i.processed !== null)
                  const totalPages = Math.ceil(eligibleRecords.length / recordsPerPage)
                  const startIdx = (previewPage - 1) * recordsPerPage
                  const endIdx = Math.min(startIdx + recordsPerPage, eligibleRecords.length)

                  return (
                    <div className="flex items-center gap-4">
                      <div className="flex items-center gap-2">
                        <span className="text-sm text-slate-400">Show</span>
                        <select
                          value={recordsPerPage}
                          onChange={(e) => {
                            setRecordsPerPage(Number(e.target.value))
                            setPreviewPage(1)
                          }}
                          className="h-8 px-2 text-sm border border-slate-600 rounded-md bg-slate-700 text-white"
                        >
                          <option value={10}>10</option>
                          <option value={20}>20</option>
                          <option value={30}>30</option>
                          <option value={50}>50</option>
                        </select>
                      </div>
                      <span className="text-sm text-slate-300">
                        {startIdx + 1}-{endIdx} of {eligibleRecords.length}
                      </span>
                      {totalPages > 1 && (
                        <div className="flex items-center gap-1">
                          <Button
                            size="sm"
                            variant="outline"
                            onClick={() => setPreviewPage(p => Math.max(1, p - 1))}
                            disabled={previewPage === 1}
                            className="h-8 w-8 p-0 border-slate-600 bg-slate-700 text-white hover:bg-slate-600 disabled:opacity-50"
                          >
                            <ChevronLeft className="w-4 h-4" />
                          </Button>
                          <span className="text-sm text-slate-300 px-2 min-w-[60px] text-center">
                            {previewPage} / {totalPages}
                          </span>
                          <Button
                            size="sm"
                            variant="outline"
                            onClick={() => setPreviewPage(p => Math.min(totalPages, p + 1))}
                            disabled={previewPage === totalPages}
                            className="h-8 w-8 p-0 border-slate-600 bg-slate-700 text-white hover:bg-slate-600 disabled:opacity-50"
                          >
                            <ChevronRight className="w-4 h-4" />
                          </Button>
                        </div>
                      )}
                    </div>
                  )
                })()}
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-sm table-fixed">
                  <thead>
                    <tr className="bg-slate-700/50 border-b border-slate-600/50">
                      <th className="px-3 py-3 text-left font-semibold text-slate-300 w-[50px]">#</th>
                      <th className="px-3 py-3 text-left font-semibold text-slate-300 w-[60px]">Title</th>
                      <th className="px-3 py-3 text-left font-semibold text-slate-300 w-[120px]">First Name</th>
                      <th className="px-3 py-3 text-left font-semibold text-slate-300 w-[120px]">Last Name</th>
                      <th className="px-3 py-3 text-left font-semibold text-slate-300 w-[200px]">Address</th>
                      <th className="px-3 py-3 text-left font-semibold text-slate-300 w-[90px]">Postcode</th>
                      <th className="px-3 py-3 text-left font-semibold text-slate-300 w-[100px]">Date</th>
                      <th className="px-3 py-3 text-right font-semibold text-slate-300 w-[90px]">Amount</th>
                      <th className="px-3 py-3 text-center font-semibold text-slate-300 w-[80px]">Status</th>
                    </tr>
                  </thead>
                  <tbody>
                    {(() => {
                      const eligibleRecords = preview.beforeAfter.filter(item => item.processed !== null)
                      const startIdx = (previewPage - 1) * recordsPerPage
                      const endIdx = startIdx + recordsPerPage
                      const pageRecords = eligibleRecords.slice(startIdx, endIdx)

                      return pageRecords.map((item, idx) => {
                        const actualIdx = startIdx + idx
                        return (
                          <tr key={`record - ${actualIdx} `} className={idx % 2 === 0 ? 'bg-slate-800/30' : 'bg-slate-800/50'}>
                            <td className="px-3 py-2.5 text-slate-400 font-medium">{actualIdx + 1}</td>
                            <td className="px-3 py-2.5 text-slate-200">{item.processed.Title}</td>
                            <td className="px-3 py-2.5 text-slate-200 truncate">{item.processed['First Name']}</td>
                            <td className="px-3 py-2.5 text-slate-200 truncate">{item.processed['Last Name']}</td>
                            <td className="px-3 py-2.5 text-slate-200 truncate">{item.processed.Address || '—'}</td>
                            <td className="px-3 py-2.5 text-slate-200 font-mono text-xs">{item.processed.Postcode || '—'}</td>
                            <td className="px-3 py-2.5 text-slate-200">{item.processed['Donation Date'] || '—'}</td>
                            <td className="px-3 py-2.5 text-right text-slate-200 font-medium">£{Number(item.processed['Donation Amount'] || 0).toFixed(2)}</td>
                            <td className="px-3 py-2.5 text-center">
                              {item.isValid ? (
                                <span className="inline-flex items-center px-2 py-0.5 rounded text-xs font-medium bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">OK</span>
                              ) : (
                                <span className="inline-flex items-center px-2 py-0.5 rounded text-xs font-medium bg-amber-500/20 text-amber-300 border border-amber-500/30">Review</span>
                              )}
                            </td>
                          </tr>
                        )
                      })
                    })()}
                  </tbody>
                </table>
              </div>
              <div className="px-5 py-3 bg-slate-800/80 border-t border-slate-700/50 flex items-center gap-4 text-sm">
                <div className="flex items-center gap-2">
                  <span className="inline-flex items-center px-2 py-0.5 rounded text-xs font-medium bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">OK</span>
                  <span className="text-slate-400">Ready for HMRC</span>
                </div>
                <div className="flex items-center gap-2">
                  <span className="inline-flex items-center px-2 py-0.5 rounded text-xs font-medium bg-amber-500/20 text-amber-300 border border-amber-500/30">Review</span>
                  <span className="text-slate-400">Needs manual check</span>
                </div>
                <div className="ml-auto text-slate-400">
                  {preview.beforeAfter.filter(i => i.processed !== null).length} records • {preview.beforeAfter.filter(i => i.processed === null).length} filtered out
                </div>
              </div>
            </div>
          </>
        )}

        {/* Error */}
        {error && (
          <div className="bg-rose-500/10 backdrop-blur-sm border-2 border-rose-500/30 rounded-2xl p-5">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-rose-500/20 flex items-center justify-center">
                <AlertCircle className="w-5 h-5 text-rose-400" />
              </div>
              <div>
                <div className="font-semibold text-rose-300">Error</div>
                <p className="text-sm text-rose-200">{error}</p>
              </div>
            </div>
          </div>
        )}

        {/* Results */}
        {result && (
          <>
            <div className="bg-slate-800/50 backdrop-blur-sm rounded-2xl shadow-xl border border-slate-700/50 p-5">
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-4">
                  <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-violet-500 to-purple-600 text-white flex items-center justify-center font-bold shadow-lg shadow-purple-500/30">
                    3
                  </div>
                  <div>
                    <h2 className="text-lg font-semibold text-white">Processing Complete!</h2>
                    <p className="text-sm text-slate-400">Your files are ready to download</p>
                  </div>
                </div>
                <div className="flex items-center gap-3">
                  <Button
                    onClick={exportAnalyticsToPDF}
                    variant="outline"
                    size="sm"
                    className="border-purple-500/50 bg-purple-500/10 text-purple-300 hover:bg-purple-500/20 hover:text-purple-200"
                  >
                    <FileText className="w-4 h-4 mr-2" />
                    Export Report
                  </Button>
                  <Button
                    onClick={handleReset}
                    size="sm"
                    className="bg-gradient-to-r from-violet-500 to-purple-600 hover:from-violet-600 hover:to-purple-700 text-white border-0"
                  >
                    <RotateCcw className="w-4 h-4 mr-2" />
                    Process Another File
                  </Button>
                  <div className="w-12 h-12 rounded-full bg-emerald-500/20 border border-emerald-500/30 flex items-center justify-center">
                    <CheckCircle2 className="w-6 h-6 text-emerald-400" />
                  </div>
                </div>
              </div>
            </div>

            <div ref={analyticsRef} className="bg-slate-800/50 backdrop-blur-sm rounded-2xl shadow-xl border border-slate-700/50 p-6">
              <div className="grid grid-cols-5 gap-4">
                <div className="bg-slate-700/30 rounded-xl p-4">
                  <div className="text-xs font-medium text-slate-400 uppercase tracking-wide">Total</div>
                  <div className="text-2xl font-bold text-white mt-1">{result.totalRecords.toLocaleString()}</div>
                </div>
                <div className="bg-gradient-to-br from-violet-500/10 to-purple-500/10 rounded-xl border border-purple-500/30 p-4">
                  <div className="text-xs font-medium text-purple-300 uppercase tracking-wide">Eligible</div>
                  <div className="text-2xl font-bold text-purple-200 mt-1">{result.eligibleRecords.toLocaleString()}</div>
                </div>
                <div className="bg-slate-700/30 rounded-xl p-4 border border-slate-600/50">
                  <div className="text-xs font-medium text-slate-400 uppercase tracking-wide">Filtered</div>
                  <div className="text-2xl font-bold text-slate-300 mt-1">{result.filteredRecords.toLocaleString()}</div>
                </div>
                <div className="bg-gradient-to-br from-emerald-500/10 to-teal-500/10 rounded-xl border border-emerald-500/30 p-4">
                  <div className="text-xs font-medium text-emerald-300 uppercase tracking-wide">Valid</div>
                  <div className="text-2xl font-bold text-emerald-200 mt-1">{result.validRecords.toLocaleString()}</div>
                </div>
                <div className="bg-amber-500/10 rounded-xl border border-amber-500/30 p-4">
                  <div className="text-xs font-medium text-amber-300 uppercase tracking-wide">Review</div>
                  <div className="text-2xl font-bold text-amber-200 mt-1">{result.invalidRecords.toLocaleString()}</div>
                </div>
              </div>

              {/* Financial Summary Section */}
              <div className="mt-6 p-5 bg-gradient-to-br from-emerald-500/10 to-teal-500/10 rounded-xl border border-emerald-500/30">
                <div className="flex items-center gap-2 mb-4">
                  <Coins className="w-5 h-5 text-emerald-400" />
                  <span className="font-semibold text-white text-lg">Financial Summary</span>
                </div>
                <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                  <div className="bg-slate-800/50 rounded-lg p-3">
                    <div className="text-xs text-slate-400 uppercase tracking-wide mb-1">Total Reviewed</div>
                    <div className="text-2xl font-bold text-emerald-200">£{result.totalAmountReviewed.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</div>
                  </div>
                  <div className="bg-slate-800/50 rounded-lg p-3">
                    <div className="text-xs text-slate-400 uppercase tracking-wide mb-1">Total Gift Aid Value</div>
                    <div className="text-2xl font-bold text-emerald-200">£{result.totalGiftAidValue.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</div>
                  </div>
                  <div className="bg-emerald-500/20 rounded-lg p-3 border border-emerald-500/30">
                    <div className="text-xs text-emerald-300 uppercase tracking-wide mb-1">Estimated Reclaimable (25%)</div>
                    <div className="text-2xl font-bold text-emerald-100">£{(result.totalGiftAidValue * 0.25).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</div>
                  </div>
                </div>
              </div>

              {/* Data Quality Section */}
              <div className="mt-6 p-5 bg-gradient-to-br from-amber-500/10 to-orange-500/10 rounded-xl border border-amber-500/30">
                <div className="flex items-center gap-2 mb-4">
                  <ShieldCheck className="w-5 h-5 text-amber-400" />
                  <span className="font-semibold text-white text-lg">Data Quality Check</span>
                </div>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div className="bg-slate-800/50 rounded-lg p-3 flex justify-between items-center">
                    <div>
                      <div className="text-xs text-slate-400 uppercase tracking-wide mb-1">Records needing attention</div>
                      <div className="text-2xl font-bold text-amber-200">{result.invalidRecords.toLocaleString()}</div>
                    </div>
                    <div className="text-right">
                      <div className="text-xs text-slate-400 uppercase tracking-wide mb-1">Valid Records</div>
                      <div className="text-xl font-semibold text-emerald-200">{result.validRecords.toLocaleString()}</div>
                    </div>
                  </div>
                  <div className="bg-slate-800/50 rounded-lg p-3 flex justify-between items-center">
                    <div>
                      <div className="text-xs text-slate-400 uppercase tracking-wide mb-1">Compliance Rate</div>
                      <div className="text-2xl font-bold text-amber-200">
                        {result.eligibleRecords > 0 ? Math.round((result.validRecords / result.eligibleRecords) * 100) : 100}%
                      </div>
                    </div>
                    <div className="w-16 h-16 relative flex items-center justify-center">
                      <svg className="w-full h-full" viewBox="0 0 36 36">
                        <path className="text-slate-700" strokeDasharray="100, 100" d="M18 2.0845 a 15.9155 15.9155 0 0 1 0 31.831 a 15.9155 15.9155 0 0 1 0 -31.831" fill="none" strokeWidth="3" />
                        <path className="text-amber-500" strokeDasharray={`${result.eligibleRecords > 0 ? (result.validRecords / result.eligibleRecords) * 100 : 100}, 100`} d="M18 2.0845 a 15.9155 15.9155 0 0 1 0 31.831 a 15.9155 15.9155 0 0 1 0 -31.831" fill="none" strokeWidth="3" strokeLinecap="round" />
                      </svg>
                    </div>
                  </div>
                </div>
              </div>

              {/* Analytics Section */}
              <div className="mt-6 p-5 bg-gradient-to-br from-purple-500/10 to-violet-500/10 rounded-xl border border-purple-500/30">
                <div className="flex items-center gap-2 mb-4">
                  <Sparkles className="w-5 h-5 text-purple-400" />
                  <span className="font-semibold text-white text-lg">Processing Details</span>
                </div>
                <div className="grid grid-cols-2 md:grid-cols-3 gap-4">
                  <div className="bg-slate-800/50 rounded-lg p-3">
                    <div className="text-xs text-slate-400 uppercase tracking-wide mb-1">Titles Assigned</div>
                    <div className="text-2xl font-bold text-purple-200">{result.analytics.titlesFilled.toLocaleString()}</div>
                  </div>
                  <div className="bg-slate-800/50 rounded-lg p-3">
                    <div className="text-xs text-slate-400 uppercase tracking-wide mb-1">Names Fixed</div>
                    <div className="text-2xl font-bold text-purple-200">{result.analytics.namesFixed.toLocaleString()}</div>
                  </div>
                  <div className="bg-slate-800/50 rounded-lg p-3">
                    <div className="text-xs text-slate-400 uppercase tracking-wide mb-1">Names from Email</div>
                    <div className="text-2xl font-bold text-purple-200">{result.analytics.namesSplitFromEmail.toLocaleString()}</div>
                  </div>
                  <div className="bg-slate-800/50 rounded-lg p-3">
                    <div className="text-xs text-slate-400 uppercase tracking-wide mb-1">Postcodes Corrected</div>
                    <div className="text-2xl font-bold text-purple-200">{result.analytics.postcodesCorrected.toLocaleString()}</div>
                  </div>
                  <div className="bg-slate-800/50 rounded-lg p-3">
                    <div className="text-xs text-slate-400 uppercase tracking-wide mb-1">Dates Formatted</div>
                    <div className="text-2xl font-bold text-purple-200">{result.analytics.datesFormatted.toLocaleString()}</div>
                  </div>
                  <div className="bg-slate-800/50 rounded-lg p-3">
                    <div className="text-xs text-slate-400 uppercase tracking-wide mb-1">Addresses Shortened</div>
                    <div className="text-2xl font-bold text-purple-200">{result.analytics.addressesShortened.toLocaleString()}</div>
                  </div>
                </div>
                <div className="mt-4 pt-4 border-t border-purple-500/20">
                  <div className="flex items-center justify-between">
                    <span className="text-sm text-slate-300">Total Cells Modified</span>
                    <span className="text-3xl font-bold bg-gradient-to-r from-purple-300 to-violet-300 bg-clip-text text-transparent">
                      {result.analytics.totalCellsModified.toLocaleString()}
                    </span>
                  </div>
                </div>
              </div>

              {/* Download Files Section */}
              <div className="mt-6 space-y-4">
                <div className="flex items-center justify-between mb-2">
                  <div className="flex items-center gap-2">
                    <Download className="w-5 h-5 text-purple-400" />
                    <span className="font-semibold text-white text-lg">Download Files</span>
                  </div>
                  <Button
                    onClick={() => {
                      const delay = (ms: number) => new Promise(r => setTimeout(r, ms))
                      const downloadFile = (url: string, name: string) => {
                        const link = document.createElement('a')
                        link.href = url
                        link.download = name
                        link.click()
                      }
                      const downloadAll = async () => {
                        if (downloadUrls.cleanedData) {
                          downloadFile(downloadUrls.cleanedData, 'cleaned_data.xlsx')
                          await delay(300)
                        }
                        for (const file of downloadUrls.hmrcFiles) {
                          downloadFile(file.url, file.name)
                          await delay(300)
                        }
                        for (const file of downloadUrls.outputFiles) {
                          downloadFile(file.url, file.name)
                          await delay(300)
                        }
                      }
                      downloadAll()
                    }}
                    size="sm"
                    className="bg-gradient-to-r from-violet-500 to-purple-600 hover:from-violet-600 hover:to-purple-700 text-white border-0 shadow-lg shadow-purple-500/20"
                  >
                    <Download className="w-4 h-4 mr-2" />
                    Download All Files
                  </Button>
                </div>

                {/* Cleaned Data File */}
                <div className="bg-gradient-to-br from-emerald-500/10 to-teal-500/10 rounded-xl border border-emerald-500/30 p-4">
                  <div className="flex items-center justify-between mb-3">
                    <div>
                      <h4 className="font-semibold text-white mb-1">Analysis & Review Package</h4>
                      <p className="text-sm text-emerald-300">cleaned_data.xlsx — 3 sheets inside</p>
                    </div>
                    {downloadUrls.cleanedData && (
                      <Button
                        onClick={() => {
                          const link = document.createElement('a')
                          link.href = downloadUrls.cleanedData!
                          link.download = 'cleaned_data.xlsx'
                          link.click()
                        }}
                        size="sm"
                        variant="outline"
                        className="border-emerald-400/50 bg-emerald-500/10 text-emerald-200 hover:bg-emerald-500/20"
                      >
                        <Download className="w-4 h-4 mr-2" />
                        Download
                      </Button>
                    )}
                  </div>
                  <div className="grid grid-cols-2 md:grid-cols-3 gap-2 mb-3">
                    <div className="bg-slate-800/40 rounded-lg p-2.5 border border-slate-600/30">
                      <div className="text-xs text-emerald-300 font-medium mb-0.5">Correct Data</div>
                      <div className="text-xs text-slate-400">All processed records ready for HMRC submission</div>
                    </div>
                    <div className="bg-slate-800/40 rounded-lg p-2.5 border border-slate-600/30">
                      <div className="text-xs text-emerald-300 font-medium mb-0.5">Needs Review</div>
                      <div className="text-xs text-slate-400">Records with missing data for manual verification</div>
                    </div>
                    <div className="bg-slate-800/40 rounded-lg p-2.5 border border-slate-600/30">
                      <div className="text-xs text-emerald-300 font-medium mb-0.5">Auto-Fixed Records</div>
                      <div className="text-xs text-slate-400">All records with automated fixes highlighted in green</div>
                    </div>
                  </div>
                </div>

                {/* HMRC Template Files */}
                <div className="bg-gradient-to-br from-violet-500/10 to-purple-500/10 rounded-xl border border-purple-500/30 p-4">
                  <div className="flex items-center justify-between mb-3">
                    <div>
                      <h4 className="font-semibold text-white mb-1">HMRC Submission Files</h4>
                      <p className="text-sm text-purple-300">Official HMRC template with your data inserted at Row 25, Column C</p>
                    </div>
                  </div>
                  <div className="grid grid-cols-2 md:grid-cols-3 gap-2 mb-3">
                    {downloadUrls.hmrcFiles.map((file, idx) => (
                      <Button
                        key={idx}
                        onClick={() => {
                          const link = document.createElement('a')
                          link.href = file.url
                          link.download = file.name
                          link.click()
                        }}
                        size="sm"
                        variant="outline"
                        className="border-purple-400/50 bg-purple-500/10 text-purple-200 hover:bg-purple-500/20 justify-start"
                      >
                        <FileSpreadsheet className="w-4 h-4 mr-2" />
                        {file.name}
                      </Button>
                    ))}
                  </div>
                  <p className="text-xs text-purple-300/70">
                    Complete HMRC template with all original formatting preserved. Up to 1000 records per file.
                  </p>
                </div>

                {/* Output Sheets */}
                <div className="bg-gradient-to-br from-cyan-500/10 to-blue-500/10 rounded-xl border border-cyan-500/30 p-4">
                  <div className="flex items-center justify-between mb-3">
                    <div>
                      <h4 className="font-semibold text-white mb-1">Simple Output Sheets</h4>
                      <p className="text-sm text-cyan-300">Clean data split into chunks of 1000 rows for easy handling</p>
                    </div>
                  </div>
                  <div className="grid grid-cols-2 md:grid-cols-3 gap-2 mb-3">
                    {downloadUrls.outputFiles.map((file, idx) => (
                      <Button
                        key={idx}
                        onClick={() => {
                          const link = document.createElement('a')
                          link.href = file.url
                          link.download = file.name
                          link.click()
                        }}
                        size="sm"
                        variant="outline"
                        className="border-cyan-400/50 bg-cyan-500/10 text-cyan-200 hover:bg-cyan-500/20 justify-start"
                      >
                        <FileSpreadsheet className="w-4 h-4 mr-2" />
                        {file.name}
                      </Button>
                    ))}
                  </div>
                  <p className="text-xs text-cyan-300/70">
                    Simple spreadsheet format without template. Perfect for backup or further processing.
                  </p>
                </div>
              </div>

              {/* HMRC Template Preview */}
              <div className="mt-6 p-4 bg-gradient-to-br from-cyan-500/10 to-blue-500/10 rounded-xl border border-cyan-500/30">
                <div className="flex items-center gap-2 mb-3">
                  <FileSpreadsheet className="w-4 h-4 text-cyan-400" />
                  <span className="font-semibold text-white">HMRC Template Preview</span>
                  <span className="text-xs text-cyan-300 ml-auto">Sample of first 3 records</span>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-xs border-collapse">
                    <thead>
                      <tr className="bg-slate-700/50 border-b border-cyan-500/30">
                        <th className="px-2 py-2 text-left text-slate-400 font-medium w-12">Row</th>
                        <th className="px-2 py-2 text-left text-cyan-300 font-medium">Col C<br />Title</th>
                        <th className="px-2 py-2 text-left text-cyan-300 font-medium">Col D<br />First name</th>
                        <th className="px-2 py-2 text-left text-cyan-300 font-medium">Col E<br />Last name</th>
                        <th className="px-2 py-2 text-left text-cyan-300 font-medium">Col F<br />House name/number</th>
                        <th className="px-2 py-2 text-left text-cyan-300 font-medium">Col G<br />Postcode</th>
                        <th className="px-2 py-2 text-left text-slate-400 font-medium">Col H<br />Aggregated</th>
                        <th className="px-2 py-2 text-left text-slate-400 font-medium">Col I<br />Sponsored</th>
                        <th className="px-2 py-2 text-left text-cyan-300 font-medium">Col J<br />Date</th>
                        <th className="px-2 py-2 text-right text-cyan-300 font-medium">Col K<br />Amount</th>
                      </tr>
                    </thead>
                    <tbody>
                      {preview && preview.beforeAfter.slice(0, 3).filter(item => item.processed).map((item, idx) => (
                        <tr key={idx} className="border-b border-slate-700/50">
                          <td className="px-2 py-2 text-slate-400">{25 + idx}</td>
                          <td className="px-2 py-2 text-slate-200">{item.processed?.Title || ''}</td>
                          <td className="px-2 py-2 text-slate-200">{item.processed?.['First Name'] || ''}</td>
                          <td className="px-2 py-2 text-slate-200">{item.processed?.['Last Name'] || ''}</td>
                          <td className="px-2 py-2 text-slate-200">{item.processed?.Address || ''}</td>
                          <td className="px-2 py-2 text-slate-200 font-mono">{item.processed?.Postcode || ''}</td>
                          <td className="px-2 py-2 text-slate-500 italic">empty</td>
                          <td className="px-2 py-2 text-slate-500 italic">empty</td>
                          <td className="px-2 py-2 text-slate-200">{item.processed?.['Donation Date'] || ''}</td>
                          <td className="px-2 py-2 text-right text-slate-200">£{Number(item.processed?.['Donation Amount'] || 0).toFixed(2)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                <div className="mt-3 text-xs text-cyan-300/70">
                  Your data is inserted starting at Row 25, Column C in the HMRC template. Rows 1-24 contain the original HMRC form headers and instructions (preserved exactly as in template).
                </div>
              </div>
            </div>
          </>
        )}
      </div>
    </div>
  )
}
