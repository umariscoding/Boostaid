'use client'

import React, { useState, useRef, useEffect } from 'react'
import * as XLSX from 'xlsx'
import ExcelJS from 'exceljs'
import { parse, isValid, format } from 'date-fns'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { AlertCircle, CheckCircle2, ChevronRight, Columns, Download, Eye, FileSpreadsheet, FileText, FileUp, Loader2, RotateCcw } from 'lucide-react'
import ColumnMapper from './column-mapper'
import { PageHeader, Pill, StatTile } from '@/components/ui/primitives'
import { Pagination } from '@/components/ui/pagination'
import { ReviewSkeleton } from '@/components/ui/skeleton'
import NoCharityNotice from '@/components/charity/no-charity-notice'
import { Select } from '@/components/ui/select'
import { useCharity } from '@/lib/store/charity'
import Link from 'next/link'
import { sanitizeRecords } from '@/lib/sanitize'
import { extractHouse } from '@/lib/address'
import { readWorkbook } from '@/lib/read-workbook'
import SummaryStrip from './claims/summary-strip'

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
  changesApplied?: string[]
  /** Why a row needs a person to look at it ("No house number", "Partial postcode", ...). */
  reasons?: string[]
}

interface ProcessingAnalytics {
  titlesFilled: number
  namesFixed: number
  namesSplitFromEmail: number
  addressesShortened: number
  postcodesCorrected: number
  datesFormatted: number
  symbolsCleaned: number
  totalCellsModified: number
}

interface PreviewData {
  sheetNames: string[]
  totalRecords: number
  eligibleRecords: number
  beforeAfter: ProcessedRecord[]
  columns: string[]
  rawRecords: any[]
  symbolsCleaned: number
}

export default function DonationProcessor() {
  // Global charity state. uploadTargetId is derived from the sidebar selection, so it is always in sync with it.
  const { loaded: charitiesLoaded, charities, selectedId, uploadTargetId, chooseUploadTarget, refresh: refreshCharities } = useCharity()
  // The charity the currently loaded file was uploaded for (fixed when the file is chosen)
  const [fileCharityId, setFileCharityId] = useState('')
  const [file, setFile] = useState<File | null>(null)
  const [preview, setPreview] = useState<PreviewData | null>(null)
  const [columnMapping, setColumnMapping] = useState<ColumnMapping | null>(null)
  const [giftAidConfig, setGiftAidConfig] = useState<GiftAidConfig | null>(null)
  const [showMapper, setShowMapper] = useState(false)
  const [processing, setProcessing] = useState(false)
  const [progressTarget, setProgressTarget] = useState(0)
  const [processingProgress, setProcessingProgress] = useState(0)
  const [loading, setLoading] = useState(false)
  const [analysing, setAnalysing] = useState(false)
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
  const [saved, setSaved] = useState<{
    claims: { id: string; ref: string; chunkIndex: number; donationCount: number; totalAmount: number }[]
    duplicates: { count: number; claimRefs: string[] }
  } | null>(null)
  const [saveError, setSaveError] = useState<string | null>(null)
  const [recordsPerPage, setRecordsPerPage] = useState(10)
  const [fixPage, setFixPage] = useState(1)
  const [fixPageSize, setFixPageSize] = useState(25)
  const [showOutput, setShowOutput] = useState(false)
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

  const removePostcodeFromAddress = (address: any, postcode: any): string => {
    if (!address) return ''
    let addrStr = String(address).trim()

    // Only remove if postcode exists and matches the valid UK format "XXX XXXX"
    if (postcode) {
      const postcodeStr = String(postcode).trim().toUpperCase()

      // Check if postcode matches valid UK format: 1-2 letters + 1-2 digits + space + digit + 2 letters
      // e.g., "AB12 3CD", "A1A 1AA", "M1 1AE"
      const validPostcodePattern = /^[A-Z]{1,2}\d{1,2}[A-Z]?\s\d[A-Z]{2}$/

      if (validPostcodePattern.test(postcodeStr)) {
        // Remove the postcode from the address if it's there
        const escapedPostcode = postcodeStr.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
        const noSpacePostcode = postcodeStr.replace(/\s/g, '')

        // Remove with space
        addrStr = addrStr.replace(new RegExp(`\\b${escapedPostcode}\\b`, 'gi'), '').trim()
        // Also remove without space in case it's concatenated
        addrStr = addrStr.replace(new RegExp(`\\b${noSpacePostcode}\\b`, 'gi'), '').trim()
      }
    }

    // Clean up multiple spaces
    addrStr = addrStr.replace(/\s+/g, ' ').trim()

    return addrStr
  }

  const hasDuplicateWords = (address: any): boolean => {
    if (!address) return false
    const addrStr = String(address).trim()

    // Split into words, convert to lowercase for comparison, and filter out empty strings
    const words = addrStr.toLowerCase().split(/\s+/).filter(w => w.length > 0)

    // Check if any word appears more than once
    const wordSet = new Set<string>()
    for (const word of words) {
      if (wordSet.has(word)) {
        return true
      }
      wordSet.add(word)
    }

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

    // HMRC wants only the house name/number: drop the postcode, then keep just the first part
    const originalAddress = String(updated['Gift Aid Address 1'] || '')
    const postcode = String(record['Gift Aid Postal Code'] || '')
    const cleanedAddress = removePostcodeFromAddress(originalAddress, postcode)
    const house = extractHouse(cleanedAddress)

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

      // Track address shortening (street/area removed so only the house name or number is sent)
      const address = String(originalRecord['Gift Aid Address 1'] || '').trim()
      if (address && house.value && house.value !== address) {
        analytics.addressesShortened++
        analytics.totalCellsModified++
        changesApplied.push('Address')
      }
    }

    const finalFirstName = String(updated['First Name'] || '').trim()
    const finalLastName = String(updated['Last Name'] || '').trim()

    // A bare number is a valid house number; no number or house name at all goes to review
    const addressHasDuplicateWords = house.value ? hasDuplicateWords(house.value) : false

    // Accents are converted to plain letters on upload, so anything outside printable ASCII left over
    // (e.g. Arabic or Cyrillic script) can't be sent to HMRC as-is and needs a person to transliterate it
    const hasUnconvertibleChars = [finalFirstName, finalLastName, house.value ?? cleanedAddress].some(v => /[^ -~]/.test(v))

    const rawPostcode = String(record['Gift Aid Postal Code'] || '').trim()
    const reasons: string[] = []
    if (needsColoringPostcode(record['Gift Aid Postal Code'])) {
      reasons.push(!rawPostcode ? 'No postcode' : rawPostcode.replace(/\s/g, '').length < 5 ? 'Partial postcode' : 'Check postcode')
    }
    if (house.value === null) reasons.push(cleanedAddress.trim() ? 'No house number' : 'No address')
    if (addressHasDuplicateWords) reasons.push('Repeated words')
    if (hasUnconvertibleChars) reasons.push('Special characters')
    if (!record['Last Donation Date']) reasons.push('No date')
    if (!finalFirstName || !finalLastName) reasons.push('Missing name')
    const needsReview = reasons.length > 0

    const processed = {
      'Title': title,
      'First Name': String(updated['First Name'] || ''),
      'Last Name': String(updated['Last Name'] || ''),
      'Address': house.value ?? cleanedAddress,
      'Postcode': needsReview ? String(record['Gift Aid Postal Code'] || '') : formatPostcode(record['Gift Aid Postal Code']),
      'Donation Date': formatDate(record['Last Donation Date']),
      'Donation Amount': Number(record['Total Donation Amount']) || 0,
      'needsReview': needsReview
    }

    return {
      original: record,
      processed: processed,
      isValid: !needsReview,
      changesApplied: changesApplied,
      reasons
    }
  }

  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      const selectedFile = e.target.files[0]
      if (!uploadTargetId) {
        setError('Choose which charity this file is for before uploading.')
        return
      }
      setFileCharityId(uploadTargetId)
      setFile(selectedFile)
      setError(null)
      setResult(null)
      setSaved(null)
      setSaveError(null)
      setFixPage(1)
      setShowOutput(false)
      setPreviewPage(1)
      setColumnMapping(null)
      setShowMapper(false)
      setLoading(true)
      setPreview(null)

      try {
        // Let the loading skeleton paint before the (synchronous) parsing starts
        await new Promise(r => setTimeout(r, 30))
        const workbook = await readWorkbook(selectedFile)

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

        // Repair garbled characters (e.g. "â€™", "ðŸ˜‡") and strip emoji/symbols from every cell
        const sanitized = sanitizeRecords(allRecords)
        allRecords = sanitized.records

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
          rawRecords: allRecords,
          symbolsCleaned: sanitized.cellsCleaned
        })
      } catch (err) {
        setError('Error reading file: ' + (err as Error).message)
      } finally {
        setLoading(false)
      }
    }
  }

  const blankAnalytics = (): ProcessingAnalytics => ({
    titlesFilled: 0,
    namesFixed: 0,
    namesSplitFromEmail: 0,
    addressesShortened: 0,
    postcodesCorrected: 0,
    datesFormatted: 0,
    symbolsCleaned: 0,
    totalCellsModified: 0
  })

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
      setAnalysing(true)
      // Let the loading skeleton paint before the (synchronous) per-row checks run
      setTimeout(() => {
        try {
          const mappedRecords = preview.rawRecords.map(rec => applyColumnMapping(rec, mapping))
          const allProcessed = mappedRecords.map(rec => processRecord(rec, updatedConfig, blankAnalytics()))
          const eligibleCount = mappedRecords.filter(rec => isEligibleDonation(rec, updatedConfig)).length

          setPreview({
            ...preview,
            eligibleRecords: eligibleCount,
            beforeAfter: allProcessed
          })
        } catch (err) {
          setError('Error checking rows: ' + (err as Error).message)
        } finally {
          setAnalysing(false)
        }
      }, 50)
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
    // Every upload is attached to a charity (the server rejects one without)
    if (!fileCharityId) {
      setError('This upload has no charity attached. Choose a charity and upload the file again.')
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
        symbolsCleaned: preview.symbolsCleaned,
        totalCellsModified: preview.symbolsCleaned
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
        .map(r => ({ ...r.processed, Reason: (r.reasons ?? []).join(', ') }))

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

      // Save the claims (one per HMRC file) so their status can be tracked later.
      // If saving fails the files are still produced, with a warning shown on screen.
      let claimRefs: string[] = []
      setSaved(null)
      setSaveError(null)
      try {
        const res = await fetch('/api/batches', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            charityId: fileCharityId,
            filename: file.name,
            stats: {
              totalRecords: allRecords.length,
              eligibleRecords: eligibleRecords.length,
              validRecords: allProcessedForCorrect.length,
              invalidRecords: recordsNeedingReview.length
            },
            claims: chunks.map(chunk => ({
              donations: chunk.map((r: any) => ({
                title: r.Title,
                firstName: r['First Name'],
                lastName: r['Last Name'],
                address: r.Address,
                postcode: r.Postcode,
                donationDate: r['Donation Date'],
                amount: Number(r['Donation Amount']) || 0
              }))
            }))
          })
        })
        const data = await res.json().catch(() => null)
        if (!res.ok) throw new Error(data?.error || `Server responded ${res.status}`)
        claimRefs = data.claims.map((c: { ref: string }) => c.ref)
        setSaved({ claims: data.claims, duplicates: data.duplicates })
        refreshCharities().catch(() => {})
      } catch (err) {
        setSaveError((err as Error).message)
      }

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

      // File names carry the schedule number, the total donations (the template's "Total donations" box) and the claim ref
      const chunkTotal = (chunk: { 'Donation Amount'?: unknown }[]) =>
        Math.round(chunk.reduce((sum: number, r) => sum + (Number(r['Donation Amount']) || 0), 0) * 100) / 100
      const fileName = (kind: string, index: number, chunk: { 'Donation Amount'?: unknown }[]) => {
        const total = chunkTotal(chunk)
        const amount = '£' + total.toLocaleString('en-GB', {
          minimumFractionDigits: Number.isInteger(total) ? 0 : 2,
          maximumFractionDigits: 2
        })
        const ref = claimRefs[index] ? ` - ${claimRefs[index]}` : ''
        return `${kind} ${index + 1} - ${amount}${ref}.xlsx`
      }

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
          name: fileName('Gift Aid Schedule', chunkIndex, chunk),
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
          name: fileName('Output Sheet', index, chunk),
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
    // The next upload must pick its charity again (matters when viewing all charities)
    chooseUploadTarget('')
    setFileCharityId('')
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
    setSaved(null)
    setSaveError(null)
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

  // ---- Derived values for the Clean screen ----
  const eligibleItems = preview ? preview.beforeAfter.filter(i => i.processed !== null) : []
  const readyCount = eligibleItems.filter(i => i.isValid).length
  const rowsToFix = eligibleItems.filter(i => !i.isValid)
  const autoFixedCount = eligibleItems.filter(i => (i.changesApplied?.length ?? 0) > 0).length
  const money = (n: number) => '£' + n.toLocaleString('en-GB', { minimumFractionDigits: 2, maximumFractionDigits: 2 })

  const downloadFile = (url: string, name: string) => {
    const link = document.createElement('a')
    link.href = url
    link.download = name
    link.click()
  }

  const downloadAll = async () => {
    const delay = (ms: number) => new Promise(r => setTimeout(r, ms))
    if (downloadUrls.cleanedData) {
      downloadFile(downloadUrls.cleanedData, 'cleaned_data.xlsx')
      await delay(300)
    }
    for (const f of [...downloadUrls.hmrcFiles, ...downloadUrls.outputFiles]) {
      downloadFile(f.url, f.name)
      await delay(300)
    }
  }

  const noCharities = charitiesLoaded && charities.length === 0
  const canUpload = !!uploadTargetId && !loading && !processing
  // "All charities" with several charities: nothing can be uploaded until one is chosen
  const needsChoice = charitiesLoaded && charities.length > 1 && selectedId === 'all' && !uploadTargetId
  const uploadTargetName = charities.find(c => c.id === uploadTargetId)?.name
  const fileCharity = charities.find(c => c.id === fileCharityId)

  const uploadButton = noCharities ? (
    <Link href="/settings" className="inline-flex h-12 items-center rounded-full bg-ink px-6 text-base font-medium text-white hover:bg-black">
      Add a charity
    </Link>
  ) : (
    <label
      title={uploadTargetId ? undefined : 'Choose a charity first'}
      className={`inline-flex h-12 items-center rounded-full bg-ink px-6 text-base font-medium text-white ${canUpload ? 'hover:bg-black cursor-pointer' : 'opacity-50 cursor-not-allowed'}`}
    >
      Upload new file
      <Input type="file" accept=".xlsx,.xls,.csv" onChange={handleFileChange} className="hidden" disabled={!canUpload} />
    </label>
  )

  const subtitle = result
    ? 'Your claim is ready.'
    : preview
      ? showMapper
        ? 'Match your columns to the fields HMRC needs.'
        : 'Check the numbers, then save the claim.'
      : 'Upload a donations spreadsheet and get it ready for HMRC.'

  return (
    <>
      <PageHeader title="Clean" subtitle={subtitle} action={uploadButton} />

      <div className="space-y-6">
        {/* ---------- Idle: nothing uploaded yet ---------- */}
        {!preview && loading && <ReviewSkeleton message="Reading your file..." />}
        {!preview && !loading && (
          <>
            {noCharities ? (
              <NoCharityNotice />
            ) : (
              <>
            <SummaryStrip refreshKey={saved?.claims.length ?? 0} />

            {needsChoice ? (
              <div className="bg-white border border-line rounded-2xl p-8">
                <h2 className="font-display text-2xl font-semibold text-ink">Choose a charity first</h2>
                <p className="text-muted mt-1">Every upload is attached to one charity. Which one is this file for?</p>
                <div className="mt-5 grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
                  {charities.map(c => (
                    <button
                      key={c.id}
                      onClick={() => chooseUploadTarget(c.id)}
                      className="rounded-xl border border-line px-5 py-4 text-left hover:border-brand/50 hover:bg-brand-soft/40 transition-colors cursor-pointer"
                    >
                      <div className="text-ink">{c.name}</div>
                      <div className="text-sm text-muted">{c.claimCount.toLocaleString()} claim{c.claimCount === 1 ? '' : 's'}</div>
                    </button>
                  ))}
                </div>
              </div>
            ) : (
              <>
            {selectedId !== 'all' || charities.length <= 1 ? (
              // A charity is already decided (picked in the sidebar, or the only one): no dropdown needed
              uploadTargetName ? (
                <p className="text-ink">
                  Uploading for <span className="font-medium">{uploadTargetName}</span>
                </p>
              ) : null
            ) : (
              <div className="flex flex-wrap items-center gap-3">
                <label htmlFor="upload-charity" className="text-ink">Uploading for</label>
                <Select
                  id="upload-charity"
                  value={uploadTargetId}
                  onChange={(e) => chooseUploadTarget(e.target.value)}
                  disabled={!charitiesLoaded}
                  wrapperClassName="min-w-[16rem]"
                >
                  {charities.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
                </Select>
              </div>
            )}

            <label className={`block ${uploadTargetId ? '' : 'opacity-50 cursor-not-allowed'}`}>
              <div className={`border-2 border-dashed border-[#cfc9e8] bg-[#faf9ff] rounded-2xl p-12 text-center transition-colors ${uploadTargetId ? 'hover:border-brand/50 hover:bg-brand-soft/50 cursor-pointer' : 'pointer-events-none'}`}>
                <div className="w-14 h-14 mx-auto mb-4 rounded-full bg-white border border-line flex items-center justify-center">
                  <FileUp className="w-6 h-6 text-brand" />
                </div>
                <div className="font-display text-xl font-semibold text-ink">Drop your file here, or click to browse</div>
                <div className="text-muted mt-1">Excel or CSV (.xlsx, .xls, .csv)</div>
                <Input type="file" accept=".xlsx,.xls,.csv" onChange={handleFileChange} className="hidden" disabled={loading || !uploadTargetId} />
              </div>
            </label>
              </>
            )}
              </>
            )}

            {!needsChoice && (
            <div className="grid md:grid-cols-3 gap-5">
              {[
                { icon: Columns, title: '1. Map columns', text: 'Match your spreadsheet columns to the fields HMRC needs.' },
                { icon: Eye, title: '2. Check the rows', text: 'See what was fixed automatically and what needs a decision.' },
                { icon: Download, title: '3. Save the claim', text: 'Get schedules named with their totals, and track the claim.' },
              ].map(({ icon: Icon, title, text }) => (
                <div key={title} className="bg-white border border-line rounded-2xl p-6 flex gap-4">
                  <div className="w-11 h-11 rounded-full bg-brand-soft text-brand flex items-center justify-center flex-shrink-0">
                    <Icon className="w-5 h-5" />
                  </div>
                  <div>
                    <div className="font-display text-lg font-semibold text-ink">{title}</div>
                    <p className="text-muted mt-1">{text}</p>
                  </div>
                </div>
              ))}
            </div>
            )}
          </>
        )}

        {/* ---------- Map columns ---------- */}
        {showMapper && preview && (
          <ColumnMapper
            detectedColumns={preview.columns}
            sampleData={preview.rawRecords}
            onMappingComplete={handleMappingComplete}
          />
        )}

        {/* ---------- Review: tiles, file card, rows to fix ---------- */}
        {preview && !showMapper && analysing && <ReviewSkeleton message="Checking every row..." />}
        {preview && !showMapper && !analysing && preview.beforeAfter.length > 0 && !result && (
          <>
            <div className="grid grid-cols-2 lg:grid-cols-4 gap-5">
              <StatTile label="Rows uploaded" value={preview.totalRecords.toLocaleString()} caption={`${preview.sheetNames.length} sheet${preview.sheetNames.length === 1 ? '' : 's'}`} />
              <StatTile label="Ready for HMRC" tone="good" value={readyCount.toLocaleString()} caption="All checks passed" />
              <StatTile label="Auto-fixed" tone="info" value={autoFixedCount.toLocaleString()} caption="House numbers, names, characters" />
              <StatTile label="To fix" tone="bad" value={rowsToFix.length.toLocaleString()} caption="Need a decision" />
            </div>

            <div className="border-2 border-dashed border-[#cfc9e8] bg-[#faf9ff] rounded-2xl p-6">
              <div className="flex flex-wrap items-center justify-between gap-4">
                <div className="min-w-0">
                  <div className="font-display text-2xl font-semibold text-ink truncate">{file?.name}</div>
                  <div className="text-muted mt-1">
                    {fileCharity ? `${fileCharity.name} · ` : ''}{preview.sheetNames.length} sheet{preview.sheetNames.length === 1 ? '' : 's'} · {preview.totalRecords.toLocaleString()} rows ·{' '}
                    {preview.eligibleRecords.toLocaleString()} Gift Aid eligible
                    {preview.totalRecords > preview.eligibleRecords && ` (${(preview.totalRecords - preview.eligibleRecords).toLocaleString()} skipped)`}
                  </div>
                </div>
                <div className="flex gap-3">
                  <Button variant="outline" size="lg" onClick={() => setShowMapper(true)} disabled={processing}>
                    Change column mapping
                  </Button>
                  <Button size="lg" onClick={processFile} disabled={processing || readyCount === 0}>
                    {processing ? (
                      <>
                        <Loader2 className="w-4 h-4 animate-spin" />
                        Saving...
                      </>
                    ) : (
                      'Save claim & make files'
                    )}
                  </Button>
                </div>
              </div>
              {processing && (
                <div className="mt-5">
                  <div className="text-sm text-muted mb-2">Processing your data...</div>
                  <div className="w-full h-2.5 bg-white border border-line rounded-full overflow-hidden">
                    <div className="h-full bg-brand transition-all duration-500 ease-out" style={{ width: `${processingProgress}%` }} />
                  </div>
                </div>
              )}
            </div>

            {/* Rows to fix */}
            <div>
              <div className="flex flex-wrap items-baseline justify-between gap-2 mb-3">
                <h2 className="font-display text-2xl font-semibold text-ink">Rows to fix ({rowsToFix.length.toLocaleString()})</h2>
                <p className="text-muted">
                  These are left out of the HMRC files. They are listed on the &ldquo;Needs Review&rdquo; sheet of the cleaned data file.
                </p>
              </div>
              <div className="bg-white border border-line rounded-2xl overflow-hidden">
                {rowsToFix.length === 0 ? (
                  <div className="py-12 text-center text-muted">Nothing to fix. Every eligible row passed all checks.</div>
                ) : (
                  <div className="overflow-x-auto">
                    <table className="w-full text-base">
                      <thead>
                        <tr className="border-b border-line">
                          <th className="px-5 py-3.5 text-left text-sm font-medium text-muted">Reason</th>
                          <th className="px-5 py-3.5 text-left text-sm font-medium text-muted">Donor</th>
                          <th className="px-5 py-3.5 text-left text-sm font-medium text-muted">House</th>
                          <th className="px-5 py-3.5 text-left text-sm font-medium text-muted">Postcode</th>
                          <th className="px-5 py-3.5 text-right text-sm font-medium text-muted">Amount</th>
                        </tr>
                      </thead>
                      <tbody>
                        {rowsToFix.slice((fixPage - 1) * fixPageSize, fixPage * fixPageSize).map((item, idx) => (
                          <tr key={idx} className="border-b border-line last:border-0">
                            <td className="px-5 py-4">
                              <div className="flex flex-wrap gap-1.5">
                                {(item.reasons ?? []).map(r => <Pill key={r} tone="bad">{r}</Pill>)}
                              </div>
                            </td>
                            <td className="px-5 py-4 text-ink">
                              {[item.processed.Title, item.processed['First Name'], item.processed['Last Name']].filter(Boolean).join(' ') || '—'}
                            </td>
                            <td className="px-5 py-4 text-ink max-w-[220px] truncate">{item.processed.Address || '—'}</td>
                            <td className="px-5 py-4 text-ink whitespace-nowrap">{item.processed.Postcode || '—'}</td>
                            <td className="px-5 py-4 text-right tabular-nums text-ink whitespace-nowrap">{money(Number(item.processed['Donation Amount'] || 0))}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
                <Pagination
                  page={fixPage}
                  pageSize={fixPageSize}
                  total={rowsToFix.length}
                  onPage={setFixPage}
                  onPageSize={(n) => {
                    setFixPageSize(n)
                    setFixPage(1)
                  }}
                />
              </div>
            </div>

            {/* Output preview (collapsed by default) */}
            <div className="bg-white border border-line rounded-2xl overflow-hidden">
              <button
                onClick={() => setShowOutput(o => !o)}
                className="w-full px-5 py-4 flex items-center justify-between text-left cursor-pointer hover:bg-wash/60"
              >
                <div>
                  <h3 className="font-display text-lg font-semibold text-ink">Preview all rows</h3>
                  <p className="text-sm text-muted">How every eligible row will look in the output</p>
                </div>
                <ChevronRight className={`w-5 h-5 text-muted transition-transform ${showOutput ? 'rotate-90' : ''}`} />
              </button>
              {showOutput && (
                <>
                  <div className="overflow-x-auto border-t border-line">
                    <table className="w-full text-sm table-fixed">
                      <thead>
                        <tr className="border-b border-line bg-wash">
                          {['#', 'Title', 'First name', 'Last name', 'House', 'Postcode', 'Date', 'Amount', 'Status'].map((h, i) => (
                            <th key={h} className={`px-3 py-3 font-medium text-muted ${i === 7 ? 'text-right' : i === 8 ? 'text-center' : 'text-left'} ${i === 0 ? 'w-[50px]' : ''}`}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {eligibleItems.slice((previewPage - 1) * recordsPerPage, previewPage * recordsPerPage).map((item, idx) => (
                          <tr key={(previewPage - 1) * recordsPerPage + idx} className="border-b border-line last:border-0">
                            <td className="px-3 py-2.5 text-muted">{(previewPage - 1) * recordsPerPage + idx + 1}</td>
                            <td className="px-3 py-2.5 text-ink">{item.processed.Title}</td>
                            <td className="px-3 py-2.5 text-ink truncate">{item.processed['First Name']}</td>
                            <td className="px-3 py-2.5 text-ink truncate">{item.processed['Last Name']}</td>
                            <td className="px-3 py-2.5 text-ink truncate">{item.processed.Address || '\u2014'}</td>
                            <td className="px-3 py-2.5 text-ink whitespace-nowrap">{item.processed.Postcode || '\u2014'}</td>
                            <td className="px-3 py-2.5 text-ink">{item.processed['Donation Date'] || '\u2014'}</td>
                            <td className="px-3 py-2.5 text-right tabular-nums text-ink">{money(Number(item.processed['Donation Amount'] || 0))}</td>
                            <td className="px-3 py-2.5 text-center">{item.isValid ? <Pill tone="good">OK</Pill> : <Pill tone="warn">Review</Pill>}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  <Pagination
                    page={previewPage}
                    pageSize={recordsPerPage}
                    total={eligibleItems.length}
                    onPage={setPreviewPage}
                    onPageSize={(n) => {
                      setRecordsPerPage(n)
                      setPreviewPage(1)
                    }}
                  />
                </>
              )}
            </div>
          </>
        )}

        {/* ---------- Error ---------- */}
        {error && (
          <div className="rounded-2xl border border-bad/20 bg-bad-soft p-5 flex items-center gap-3">
            <AlertCircle className="w-5 h-5 text-bad flex-shrink-0" />
            <div>
              <div className="font-medium text-bad">Something went wrong</div>
              <p className="text-sm text-bad/90">{error}</p>
            </div>
          </div>
        )}

        {/* ---------- Results ---------- */}
        {result && (
          <>
            <div className="bg-white border border-line rounded-2xl p-6 flex flex-wrap items-center justify-between gap-4">
              <div className="flex items-center gap-4">
                <div className="w-12 h-12 rounded-full bg-good-soft flex items-center justify-center">
                  <CheckCircle2 className="w-6 h-6 text-good" />
                </div>
                <div>
                  <h2 className="font-display text-2xl font-semibold text-ink">Claim saved</h2>
                  <p className="text-muted">{fileCharity ? `Saved for ${fileCharity.name}. ` : ''}Your files are ready to download.</p>
                </div>
              </div>
              <div className="flex gap-3">
                <Button variant="outline" onClick={exportAnalyticsToPDF}>
                  <FileText className="w-4 h-4" /> Export report
                </Button>
                <Button variant="dark" onClick={handleReset}>
                  <RotateCcw className="w-4 h-4" /> Clean another file
                </Button>
              </div>
            </div>

            <div ref={analyticsRef} className="grid grid-cols-2 lg:grid-cols-5 gap-5">
              <StatTile label="Rows uploaded" value={result.totalRecords.toLocaleString()} />
              <StatTile label="Eligible" tone="brand" value={result.eligibleRecords.toLocaleString()} />
              <StatTile label="Skipped" value={result.filteredRecords.toLocaleString()} caption="Not Gift Aid" />
              <StatTile label="Ready for HMRC" tone="good" value={result.validRecords.toLocaleString()} />
              <StatTile label="To fix" tone="bad" value={result.invalidRecords.toLocaleString()} />
            </div>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-5">
              <StatTile label="Total donations" value={money(result.totalAmountReviewed)} caption="All eligible rows" />
              <StatTile label="In the HMRC files" value={money(result.totalGiftAidValue)} caption="Rows that passed every check" />
              <StatTile label="Gift Aid to reclaim (25%)" tone="good" value={money(result.totalGiftAidValue * 0.25)} />
            </div>

            {saveError && (
              <div className="rounded-2xl border border-bad/20 bg-bad-soft p-5 flex gap-3 text-sm text-bad">
                <AlertCircle className="w-5 h-5 flex-shrink-0" />
                <div>
                  <div className="font-medium">Claims were not saved</div>
                  Your files below are fine, but this upload will not appear in Claims or Search ({saveError}). Process the file again once the problem is fixed.
                </div>
              </div>
            )}

            {saved && (
              <div className="bg-white border border-line rounded-2xl p-6">
                <div className="flex flex-wrap items-center justify-between gap-3 mb-4">
                  <div className="flex items-baseline gap-3">
                    <h3 className="font-display text-xl font-semibold text-ink">Claims created</h3>
                    <span className="text-muted">all start as Pending</span>
                  </div>
                  <Link href="/claims" className="font-medium text-brand hover:underline">Open Claims &rarr;</Link>
                </div>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                  {saved.claims.map(c => (
                    <Link
                      key={c.id}
                      href={`/claims/${c.id}`}
                      className="flex items-center justify-between rounded-xl border border-line px-5 py-4 hover:bg-wash transition-colors"
                    >
                      <div>
                        <div className="text-ink">{c.ref}</div>
                        <div className="text-sm text-muted">{c.donationCount.toLocaleString()} donations</div>
                      </div>
                      <div className="text-right">
                        <div className="text-ink">{money(c.totalAmount)}</div>
                        <div className="text-sm text-good">{money(c.totalAmount * 0.25)} Gift Aid</div>
                      </div>
                    </Link>
                  ))}
                </div>
                {saved.duplicates.count > 0 && (
                  <div className="mt-4 p-4 bg-warn-soft rounded-xl text-sm text-warn flex gap-3">
                    <AlertCircle className="w-5 h-5 flex-shrink-0" />
                    <div>
                      <span className="font-medium">
                        {saved.duplicates.count.toLocaleString()} donation{saved.duplicates.count === 1 ? '' : 's'}
                      </span>{' '}
                      in this upload already appear in earlier claims ({saved.duplicates.claimRefs.join(', ')}). Check you are not claiming Gift Aid on them twice.
                    </div>
                  </div>
                )}
              </div>
            )}

            {/* Downloads */}
            <div className="bg-white border border-line rounded-2xl p-6 space-y-6">
              <div className="flex flex-wrap items-center justify-between gap-3">
                <h3 className="font-display text-xl font-semibold text-ink">Download files</h3>
                <Button onClick={downloadAll}><Download className="w-4 h-4" /> Download all</Button>
              </div>

              <div>
                <div className="font-medium text-ink">HMRC schedules</div>
                <p className="text-sm text-muted mb-3">The official HMRC template with your data from row 25. Up to 1,000 donations per file; the amount in each name is its total.</p>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-2">
                  {downloadUrls.hmrcFiles.map((f, i) => (
                    <Button key={i} variant="outline" className="justify-start h-auto py-3 whitespace-normal text-left" onClick={() => downloadFile(f.url, f.name)}>
                      <FileSpreadsheet className="w-4 h-4 flex-shrink-0 text-good" />
                      {f.name}
                    </Button>
                  ))}
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div>
                  <div className="font-medium text-ink">Review package</div>
                  <p className="text-sm text-muted mb-3">cleaned_data.xlsx: correct data, rows to fix (with reasons), skipped rows and every auto-fix.</p>
                  {downloadUrls.cleanedData && (
                    <Button variant="outline" className="justify-start" onClick={() => downloadFile(downloadUrls.cleanedData!, 'cleaned_data.xlsx')}>
                      <FileSpreadsheet className="w-4 h-4 text-info" /> cleaned_data.xlsx
                    </Button>
                  )}
                </div>
                <div>
                  <div className="font-medium text-ink">Simple output sheets</div>
                  <p className="text-sm text-muted mb-3">Plain spreadsheets of the clean data, for backup or further work.</p>
                  <div className="flex flex-col gap-2">
                    {downloadUrls.outputFiles.map((f, i) => (
                      <Button key={i} variant="outline" className="justify-start h-auto py-3 whitespace-normal text-left" onClick={() => downloadFile(f.url, f.name)}>
                        <FileSpreadsheet className="w-4 h-4 flex-shrink-0 text-info" />
                        {f.name}
                      </Button>
                    ))}
                  </div>
                </div>
              </div>
            </div>

            {/* What was fixed */}
            <div className="bg-white border border-line rounded-2xl p-6">
              <h3 className="font-display text-xl font-semibold text-ink mb-4">What we fixed</h3>
              <div className="grid grid-cols-2 md:grid-cols-4 gap-x-8 gap-y-5">
                {[
                  ['Titles assigned', result.analytics.titlesFilled],
                  ['Names fixed', result.analytics.namesFixed],
                  ['Names from email', result.analytics.namesSplitFromEmail],
                  ['Postcodes corrected', result.analytics.postcodesCorrected],
                  ['Dates formatted', result.analytics.datesFormatted],
                  ['Addresses shortened', result.analytics.addressesShortened],
                  ['Symbols cleaned', result.analytics.symbolsCleaned],
                  ['Cells changed in total', result.analytics.totalCellsModified],
                ].map(([label, value]) => (
                  <div key={label as string}>
                    <div className="font-display text-2xl font-bold text-ink">{(value as number).toLocaleString()}</div>
                    <div className="text-sm text-muted">{label}</div>
                  </div>
                ))}
              </div>
            </div>

            {/* HMRC template preview */}
            <div className="bg-white border border-line rounded-2xl overflow-hidden">
              <div className="px-6 py-4 border-b border-line flex items-baseline justify-between">
                <h3 className="font-display text-xl font-semibold text-ink">HMRC schedule preview</h3>
                <span className="text-sm text-muted">first 3 rows, starting at row 25</span>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="border-b border-line bg-wash text-muted">
                      {['Row', 'Title', 'First name', 'Last name', 'House name or number', 'Postcode', 'Aggregated', 'Sponsored', 'Date', 'Amount'].map((h, i) => (
                        <th key={h} className={`px-4 py-3 font-medium ${i === 9 ? 'text-right' : 'text-left'}`}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {preview && preview.beforeAfter.filter(i => i.processed && i.isValid).slice(0, 3).map((item, idx) => (
                      <tr key={idx} className="border-b border-line last:border-0">
                        <td className="px-4 py-3 text-muted">{25 + idx}</td>
                        <td className="px-4 py-3 text-ink">{item.processed?.Title || ''}</td>
                        <td className="px-4 py-3 text-ink">{item.processed?.['First Name'] || ''}</td>
                        <td className="px-4 py-3 text-ink">{item.processed?.['Last Name'] || ''}</td>
                        <td className="px-4 py-3 text-ink">{item.processed?.Address || ''}</td>
                        <td className="px-4 py-3 text-ink">{item.processed?.Postcode || ''}</td>
                        <td className="px-4 py-3 text-muted italic">empty</td>
                        <td className="px-4 py-3 text-muted italic">empty</td>
                        <td className="px-4 py-3 text-ink">{item.processed?.['Donation Date'] || ''}</td>
                        <td className="px-4 py-3 text-right tabular-nums text-ink">{money(Number(item.processed?.['Donation Amount'] || 0))}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </>
        )}
      </div>
    </>
  )
}
