'use client'

import { useState } from 'react'
import { Button } from './ui/button'
import { ArrowRight, Check, RotateCcw, GripVertical, X } from 'lucide-react'

interface ColumnMapping {
  [key: string]: string | string[] | null
}

interface GiftAidConfig {
  column: string | null
  valuesToKeep: string[]
}

interface ColumnMapperProps {
  detectedColumns: string[]
  sampleData: any[]
  onMappingComplete: (mapping: ColumnMapping, giftAidConfig: GiftAidConfig) => void
}

const REQUIRED_COLUMNS = [
  { key: 'Title', required: false, allowMultiple: false },
  { key: 'First Name', required: true, allowMultiple: false },
  { key: 'Last Name', required: true, allowMultiple: false },
  { key: 'Email', required: false, allowMultiple: false },
  { key: 'Address', required: true, allowMultiple: true },
  { key: 'Postcode', required: true, allowMultiple: false },
  { key: 'Date', required: true, allowMultiple: false },
  { key: 'Amount', required: true, allowMultiple: false },
  { key: 'Gift Aid', required: true, allowMultiple: false },
]

function ColumnMapper({ detectedColumns, sampleData, onMappingComplete }: ColumnMapperProps) {
  const [mapping, setMapping] = useState<ColumnMapping>({})
  const [draggedColumn, setDraggedColumn] = useState<string | null>(null)
  const [dragOverTarget, setDragOverTarget] = useState<string | null>(null)
  const [giftAidValues, setGiftAidValues] = useState<string[]>([])
  const [selectedGiftAidValues, setSelectedGiftAidValues] = useState<string[]>([])

  const autoDetectMapping = (): ColumnMapping => {
    const autoMapping: ColumnMapping = {}
    
    const patterns = {
      'Title': ['title', 'donortitle', 'prefix'],
      'First Name': ['firstname', 'first_name', 'donorfirstname', 'forename', 'givenname'],
      'Last Name': ['lastname', 'last_name', 'surname', 'donorsurname', 'familyname'],
      'Email': ['email', 'donoremail', 'emailaddress'],
      'Address': ['address', 'addressline', 'donoraddress', 'street'],
      'Postcode': ['postcode', 'postal_code', 'donoraddresspostcode', 'zipcode', 'zip'],
      'Date': ['date', 'donationdate', 'transactiondate', 'paymentdate'],
      'Amount': ['amount', 'donationamount', 'value', 'donationgrossamount'],
      'Gift Aid': ['giftaid', 'taxeffective', 'donationtaxeffective', 'giftaidstatus', 'tax_eligible'],
    }

    REQUIRED_COLUMNS.forEach(({ key, allowMultiple }) => {
      const possiblePatterns = patterns[key as keyof typeof patterns] || []
      
      if (allowMultiple && key === 'Address') {
        const matches = detectedColumns.filter(col => {
          const colLower = col.toLowerCase().replace(/[^a-z0-9]/g, '')
          // "Email Address" is not a postal address
          if (colLower.includes('email')) return false
          return possiblePatterns.some(pattern => colLower.includes(pattern))
        })
        if (matches.length > 0) {
          autoMapping[key] = matches.slice(0, 3)
        }
      } else {
        const match = detectedColumns.find(col => {
          const colLower = col.toLowerCase().replace(/[^a-z0-9]/g, '')
          return possiblePatterns.some(pattern => colLower.includes(pattern))
        })
        if (match) {
          autoMapping[key] = match
          
          if (key === 'Gift Aid') {
            const uniqueValues = new Set<string>()
            sampleData.forEach(row => {
              const value = String(row[match] || '').trim()
              if (value) uniqueValues.add(value)
            })
            const values = Array.from(uniqueValues)
            setGiftAidValues(values)
            
            const autoSelect = values.filter(v => {
              const lower = v.toLowerCase()
              return lower.includes('yes') || 
                     (lower.includes('effective') && !lower.includes('non')) ||
                     lower === 'y' || 
                     lower === 'true'
            })
            setSelectedGiftAidValues(autoSelect.length > 0 ? autoSelect : values)
          }
        }
      }
    })

    return autoMapping
  }

  const handleAutoMap = () => {
    const autoMapping = autoDetectMapping()
    setMapping(autoMapping)
  }

  const handleDragStart = (column: string) => {
    setDraggedColumn(column)
  }

  const handleDragEnd = () => {
    setDraggedColumn(null)
    setDragOverTarget(null)
  }

  const handleDragOver = (e: React.DragEvent, targetColumn: string) => {
    e.preventDefault()
    if (dragOverTarget !== targetColumn) {
      setDragOverTarget(targetColumn)
    }
  }

  const handleDragLeave = (e: React.DragEvent) => {
    // Only reset if we're actually leaving the card, not moving to a child element
    if (!e.currentTarget.contains(e.relatedTarget as Node)) {
      setDragOverTarget(null)
    }
  }

  const handleDrop = (targetColumn: string) => {
    if (!draggedColumn) return
    
    const col = REQUIRED_COLUMNS.find(c => c.key === targetColumn)
    
    if (col?.allowMultiple) {
      const current = mapping[targetColumn]
      const currentArray = Array.isArray(current) ? current : current ? [current] : []
      if (!currentArray.includes(draggedColumn)) {
        setMapping(prev => ({
          ...prev,
          [targetColumn]: [...currentArray, draggedColumn]
        }))
      }
    } else {
      setMapping(prev => ({
        ...prev,
        [targetColumn]: draggedColumn
      }))
      
      if (targetColumn === 'Gift Aid') {
        const uniqueValues = new Set<string>()
        sampleData.forEach(row => {
          const value = String(row[draggedColumn] || '').trim()
          if (value) uniqueValues.add(value)
        })
        const values = Array.from(uniqueValues)
        setGiftAidValues(values)
        
        const autoSelect = values.filter(v => {
          const lower = v.toLowerCase()
          return lower.includes('yes') || 
                 (lower.includes('effective') && !lower.includes('non')) ||
                 lower === 'y' || 
                 lower === 'true'
        })
        setSelectedGiftAidValues(autoSelect.length > 0 ? autoSelect : values)
      }
    }
    
    setDraggedColumn(null)
    setDragOverTarget(null)
  }

  const handleRemoveMapping = (targetCol: string, sourceCol?: string) => {
    const col = REQUIRED_COLUMNS.find(c => c.key === targetCol)
    
    if (col?.allowMultiple && sourceCol) {
      const current = mapping[targetCol]
      if (Array.isArray(current)) {
        const filtered = current.filter(c => c !== sourceCol)
        setMapping(prev => ({
          ...prev,
          [targetCol]: filtered.length > 0 ? filtered : null
        }))
      }
    } else {
      setMapping(prev => {
        const newMapping = { ...prev }
        delete newMapping[targetCol]
        return newMapping
      })
      
      if (targetCol === 'Gift Aid') {
        setGiftAidValues([])
        setSelectedGiftAidValues([])
      }
    }
  }

  const isComplete = () => {
    const allRequiredMapped = REQUIRED_COLUMNS
      .filter(col => col.required)
      .every(col => {
        const mapped = mapping[col.key]
        return mapped && (Array.isArray(mapped) ? mapped.length > 0 : true)
      })
    
    const giftAidConfigured = mapping['Gift Aid'] && selectedGiftAidValues.length > 0
    
    return allRequiredMapped && giftAidConfigured
  }

  const handleComplete = () => {
    const giftAidConfig: GiftAidConfig = {
      column: typeof mapping['Gift Aid'] === 'string' ? mapping['Gift Aid'] : null,
      valuesToKeep: selectedGiftAidValues
    }
    onMappingComplete(mapping, giftAidConfig)
  }

  const getUsedColumns = () => {
    const used = new Set<string>()
    Object.values(mapping).forEach(val => {
      if (Array.isArray(val)) {
        val.forEach(v => used.add(v))
      } else if (val) {
        used.add(val)
      }
    })
    return used
  }

  const usedColumns = getUsedColumns()

  const requiredCount = REQUIRED_COLUMNS.filter(c => c.required).length
  const mappedCount = Object.keys(mapping).filter(k => REQUIRED_COLUMNS.find(c => c.key === k && c.required) && mapping[k]).length

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="bg-white border border-line rounded-2xl p-6">
        <div className="flex flex-wrap items-center justify-between gap-4">
          <div>
            <h2 className="font-display text-2xl font-semibold text-ink">Map your columns</h2>
            <p className="text-muted mt-1">Drag column headers from the table below onto the matching boxes, or let us detect them.</p>
          </div>
          <div className="flex gap-3">
            <Button onClick={handleAutoMap} variant="outline">
              <RotateCcw className="w-4 h-4" />
              Auto-Detect
            </Button>
            <Button onClick={handleComplete} disabled={!isComplete()}>
              <Check className="w-4 h-4" />
              Continue ({mappedCount}/{requiredCount})
            </Button>
          </div>
        </div>
      </div>

      {/* Mapping Boxes */}
      <div className="grid gap-4" style={{ gridTemplateColumns: 'repeat(20, 1fr)' }}>
        {REQUIRED_COLUMNS.map((col, index) => {
          const mappedValue = mapping[col.key]
          const isMapped = mappedValue && (Array.isArray(mappedValue) ? mappedValue.length > 0 : true)
          const mappedArray = Array.isArray(mappedValue) ? mappedValue : mappedValue ? [mappedValue] : []
          const isHovering = dragOverTarget === col.key && draggedColumn

          return (
            <div
              key={col.key}
              onDragOver={(e) => handleDragOver(e, col.key)}
              onDragLeave={handleDragLeave}
              onDrop={() => handleDrop(col.key)}
              style={{ gridColumn: index < 5 ? 'span 4' : 'span 5' }}
              className={`p-4 rounded-2xl border-2 border-dashed min-h-[110px] transition-all duration-200 ease-out ${
                isHovering
                  ? 'border-brand bg-brand-soft scale-[1.02] ring-2 ring-brand/20'
                  : isMapped
                    ? 'border-brand/40 bg-brand-soft/60'
                    : draggedColumn
                      ? 'border-brand/40 bg-wash'
                      : 'border-line bg-white hover:border-brand/30'
              }`}
            >
              <div className="text-sm font-medium text-ink mb-2">
                {col.key}
                {col.required && <span className="text-bad ml-1">*</span>}
              </div>

              {isMapped ? (
                <div className="space-y-2">
                  {mappedArray.map((source, idx) => (
                    <div key={idx} className="flex items-center gap-1 text-xs bg-white rounded-lg px-2.5 py-1.5 border border-brand/20">
                      <span className="flex-1 font-medium text-brand truncate" title={source}>{source}</span>
                      <button onClick={() => handleRemoveMapping(col.key, source)} className="flex-shrink-0 cursor-pointer">
                        <X className="w-3 h-3 text-muted hover:text-bad" />
                      </button>
                    </div>
                  ))}

                  {col.key === 'Gift Aid' && giftAidValues.length > 0 && (
                    <div className="mt-2 pt-2 border-t border-brand/20">
                      <div className="text-xs text-muted mb-1.5">Keep values:</div>
                      <div className="flex flex-wrap gap-1">
                        {giftAidValues.map(value => (
                          <label
                            key={value}
                            className={`px-2.5 py-1 rounded-full text-xs font-medium cursor-pointer transition-colors ${
                              selectedGiftAidValues.includes(value)
                                ? 'bg-brand text-white'
                                : 'bg-white border border-line text-muted hover:text-ink'
                            }`}
                          >
                            <input
                              type="checkbox"
                              checked={selectedGiftAidValues.includes(value)}
                              onChange={(e) => {
                                if (e.target.checked) {
                                  setSelectedGiftAidValues([...selectedGiftAidValues, value])
                                } else {
                                  setSelectedGiftAidValues(selectedGiftAidValues.filter(v => v !== value))
                                }
                              }}
                              className="hidden"
                            />
                            {value}
                          </label>
                        ))}
                      </div>
                    </div>
                  )}
                </div>
              ) : (
                <div className="text-xs text-muted mt-2">Drop column here</div>
              )}
            </div>
          )
        })}
      </div>

      {/* Data Table */}
      <div className="bg-white border border-line rounded-2xl overflow-hidden">
        <div className="px-5 py-3.5 border-b border-line flex items-center justify-between">
          <span className="font-display text-lg font-semibold text-ink">Your data</span>
          <span className="text-sm text-muted">{detectedColumns.length} columns, first 10 rows</span>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="border-b border-line">
                {detectedColumns.map((col) => {
                  const isUsed = usedColumns.has(col)
                  return (
                    <th
                      key={col}
                      draggable={!isUsed}
                      onDragStart={() => handleDragStart(col)}
                      onDragEnd={handleDragEnd}
                      className={`px-3 py-3 text-left font-medium whitespace-nowrap transition-colors ${
                        isUsed
                          ? 'bg-brand-soft text-brand cursor-not-allowed'
                          : draggedColumn === col
                            ? 'bg-brand-soft text-brand cursor-grabbing'
                            : 'bg-wash text-ink cursor-grab hover:bg-brand-soft/60'
                      }`}
                    >
                      <div className="flex items-center gap-1.5">
                        {!isUsed && <GripVertical className="w-3.5 h-3.5 text-muted" />}
                        <span className="text-xs" title={col}>{col}</span>
                        {isUsed && <Check className="w-3.5 h-3.5 text-brand" />}
                      </div>
                    </th>
                  )
                })}
              </tr>
            </thead>
            <tbody>
              {sampleData.slice(0, 10).map((row, idx) => (
                <tr key={idx} className="border-b border-line last:border-0">
                  {detectedColumns.map((col) => {
                    const value = row[col]
                    let displayValue = value
                    if (value instanceof Date) {
                      displayValue = value.toLocaleDateString('en-GB')
                    } else if (typeof value === 'number' && col.toLowerCase().includes('date')) {
                      const excelEpoch = new Date(1899, 11, 30)
                      const date = new Date(excelEpoch.getTime() + value * 86400000)
                      displayValue = date.toLocaleDateString('en-GB')
                    } else {
                      displayValue = String(value || '')
                    }

                    return (
                      <td key={col} className="px-3 py-2 text-ink text-xs max-w-[220px] truncate">
                        {displayValue}
                      </td>
                    )
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {/* Status Bar */}
      <div className={`flex items-center justify-between p-4 rounded-2xl border ${
        isComplete() ? 'bg-brand-soft border-brand/30' : 'bg-wash border-line'
      }`}>
        <div className="flex items-center gap-2">
          {isComplete() ? (
            <Check className="w-5 h-5 text-brand" />
          ) : (
            <ArrowRight className="w-5 h-5 text-muted" />
          )}
          <span className={`font-medium ${isComplete() ? 'text-brand' : 'text-ink'}`}>
            {isComplete() ? 'Ready to continue' : 'Map all required fields to continue'}
          </span>
        </div>
        <span className="text-sm text-muted">{mappedCount} of {requiredCount} required</span>
      </div>
    </div>
  )
}

export default ColumnMapper
