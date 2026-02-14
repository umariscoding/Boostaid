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

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="bg-slate-800/50 backdrop-blur-sm rounded-2xl shadow-xl border border-slate-700/50 p-5">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-4">
            <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-violet-500 to-purple-600 text-white flex items-center justify-center font-bold shadow-lg shadow-purple-500/30">1</div>
            <div>
              <h2 className="text-lg font-semibold text-white">Map Your Columns</h2>
              <p className="text-sm text-slate-400">Drag column headers onto the target boxes below</p>
            </div>
          </div>
          <div className="flex gap-3">
            <Button onClick={handleAutoMap} variant="outline" size="sm" className="h-9 border-slate-600 bg-slate-700 text-white hover:bg-slate-600">
              <RotateCcw className="w-4 h-4 mr-2" />
              Auto-Detect
            </Button>
            <Button onClick={handleComplete} disabled={!isComplete()} size="sm" className="h-9 bg-gradient-to-r from-violet-500 to-purple-600 hover:from-violet-600 hover:to-purple-700 text-white border-0">
              <Check className="w-4 h-4 mr-2" />
              Continue ({Object.keys(mapping).filter(k => REQUIRED_COLUMNS.find(c => c.key === k && c.required) && mapping[k]).length}/{REQUIRED_COLUMNS.filter(c => c.required).length})
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
              className={`p-4 rounded-xl border-2 border-dashed min-h-[100px] transition-all duration-200 ease-out ${
                isHovering
                  ? 'border-purple-400 bg-purple-500/20 scale-[1.02] shadow-lg shadow-purple-500/20 ring-2 ring-purple-400/50'
                  : isMapped 
                    ? 'border-purple-500/50 bg-purple-500/10' 
                    : draggedColumn
                      ? 'border-purple-400/40 bg-purple-500/5'
                      : 'border-slate-600 bg-slate-800/50 hover:border-slate-500'
              }`}
            >
              <div className="text-sm font-semibold text-white mb-2">
                {col.key}
                {col.required && <span className="text-rose-400 ml-1">*</span>}
              </div>
              
              {isMapped ? (
                <div className="space-y-2">
                  {mappedArray.map((source, idx) => (
                    <div key={idx} className="flex items-center gap-1 text-xs bg-slate-700/50 rounded-lg px-2.5 py-1.5 border border-purple-500/30">
                      <span className="flex-1 font-medium text-purple-200 truncate" title={source}>{source}</span>
                      <button onClick={() => handleRemoveMapping(col.key, source)} className="flex-shrink-0">
                        <X className="w-3 h-3 text-slate-400 hover:text-rose-400" />
                      </button>
                    </div>
                  ))}
                  
                  {col.key === 'Gift Aid' && giftAidValues.length > 0 && (
                    <div className="mt-2 pt-2 border-t border-slate-600">
                      <div className="text-xs text-slate-400 mb-1.5">Keep values:</div>
                      <div className="flex flex-wrap gap-1">
                        {giftAidValues.map(value => (
                          <label
                            key={value}
                            className={`px-2 py-1 rounded-md text-xs cursor-pointer transition-colors ${
                              selectedGiftAidValues.includes(value)
                                ? 'bg-gradient-to-r from-violet-500 to-purple-600 text-white'
                                : 'bg-slate-700 text-slate-300 hover:bg-slate-600'
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
                <div className="text-xs text-slate-500 mt-2">Drop column here</div>
              )}
            </div>
          )
        })}
      </div>

      {/* Data Table */}
      <div className="bg-slate-800/50 backdrop-blur-sm rounded-2xl shadow-xl border border-slate-700/50 overflow-hidden">
        <div className="px-5 py-3 bg-slate-800/80 border-b border-slate-700/50 flex items-center justify-between">
          <span className="font-semibold text-white">Your Data</span>
          <span className="text-sm text-slate-400">{detectedColumns.length} columns preview</span>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="bg-slate-700/50 border-b border-slate-600/50">
                {detectedColumns.map((col) => {
                  const isUsed = usedColumns.has(col)
                  return (
                    <th
                      key={col}
                      draggable={!isUsed}
                      onDragStart={() => handleDragStart(col)}
                      onDragEnd={handleDragEnd}
                      className={`px-3 py-2.5 text-left font-semibold whitespace-nowrap transition-colors ${
                        isUsed 
                          ? 'bg-purple-500/20 text-purple-200 cursor-not-allowed' 
                          : draggedColumn === col
                            ? 'bg-purple-500/30 text-purple-100 cursor-grabbing'
                            : 'bg-slate-700/50 text-slate-300 cursor-grab hover:bg-slate-600/50'
                      }`}
                    >
                      <div className="flex items-center gap-1.5">
                        {!isUsed && <GripVertical className="w-3.5 h-3.5 text-slate-500" />}
                        <span className="text-xs" title={col}>{col}</span>
                        {isUsed && <Check className="w-3.5 h-3.5 text-purple-400" />}
                      </div>
                    </th>
                  )
                })}
              </tr>
            </thead>
            <tbody>
              {sampleData.slice(0, 10).map((row, idx) => (
                <tr key={idx} className={idx % 2 === 0 ? 'bg-slate-800/30' : 'bg-slate-800/50'}>
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
                      <td key={col} className="px-3 py-2 text-slate-300 text-xs">
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
      <div className={`flex items-center justify-between p-4 rounded-xl border-2 ${
        isComplete() 
          ? 'bg-purple-500/10 border-purple-500/30' 
          : 'bg-slate-800/50 border-slate-600'
      }`}>
        <div className="flex items-center gap-2">
          {isComplete() ? (
            <Check className="w-5 h-5 text-purple-400" />
          ) : (
            <ArrowRight className="w-5 h-5 text-slate-400" />
          )}
          <span className={`font-semibold ${isComplete() ? 'text-purple-200' : 'text-slate-300'}`}>
            {isComplete() ? 'Ready to continue' : 'Map all required fields to continue'}
          </span>
        </div>
        <span className={`text-sm ${isComplete() ? 'text-purple-300' : 'text-slate-400'}`}>
          {Object.keys(mapping).filter(k => REQUIRED_COLUMNS.find(c => c.key === k && c.required) && mapping[k]).length} of {REQUIRED_COLUMNS.filter(c => c.required).length} required
        </span>
      </div>
    </div>
  )
}

export default ColumnMapper
