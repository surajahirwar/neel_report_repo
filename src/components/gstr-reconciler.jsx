import { useState, useEffect } from "react"
import {
  Upload,
  FileSpreadsheet,
  AlertCircle,
  CheckCircle,
  XCircle,
  Download,
  RefreshCw,
  Loader2,
  HelpCircle,
  AlertTriangle,
} from "lucide-react"

// Helper to clean invoice numbers for better matching (removes special chars, spaces)
const cleanString = (str) => {
  if (!str) return ""
  return str
    .toString()
    .replace(/[^a-zA-Z0-9]/g, "")
    .toUpperCase()
}

// Helper to check string similarity (inclusion)
const isSimilarString = (str1, str2) => {
  if (!str1 || !str2) return false
  const s1 = cleanString(str1)
  const s2 = cleanString(str2)
  // Check if one contains the other and length difference is small (e.g., typo of 1-2 chars)
  const lengthDiff = Math.abs(s1.length - s2.length)
  return (s1.includes(s2) || s2.includes(s1)) && lengthDiff <= 2
}

// Helper to parse numeric values safely
const parseNumber = (val) => {
  if (typeof val === "number") return val
  if (!val) return 0
  // Remove currency symbols and commas
  const clean = val.toString().replace(/[₹,]/g, "")
  const num = Number.parseFloat(clean)
  return isNaN(num) ? 0 : num
}

export default function GstrReconciler() {
  const [step, setStep] = useState(1)
  const [isLoading, setIsLoading] = useState(false)
  const [error, setError] = useState("")
  const [reconciliationData, setReconciliationData] = useState([])
  const [summary, setSummary] = useState({
    total: 0,
    matched: 0,
    mismatch: 0,
    billMismatch: 0,
    probableMatch: 0,
    missingInGstr: 0,
    missingInSoft: 0,
  })
  const [libLoaded, setLibLoaded] = useState(false)

  // Load SheetJS library dynamically
  useEffect(() => {
    const script = document.createElement("script")
    script.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"
    script.async = true
    script.onload = () => setLibLoaded(true)
    document.body.appendChild(script)
    return () => {
      try {
        document.body.removeChild(script)
      } catch {
        // Script already removed
      }
    }
  }, [])

  // Generic parser for array of arrays (from Excel)
  const parseSheetData = (rows, type) => {
    if (!rows || rows.length < 2) return []

    // Find header row (assuming it's the first non-empty row)
    let headerRowIndex = 0
    for (let i = 0; i < rows.length; i++) {
      if (rows[i].length > 1) {
        headerRowIndex = i
        break
      }
    }

    const headers = rows[headerRowIndex].map((h) => (h ? h.toString().trim() : ""))
    const data = []

    // Dynamic column mapping
    const map = {}
    headers.forEach((h, i) => {
      const lower = h.toLowerCase()
      if (type === "GSTR") {
        if (lower.includes("gstin")) map.gstin = i
        else if (lower.includes("invoice number") || lower === "invoice no") map.inv = i
        else if (lower.includes("taxable value")) map.taxable = i
        else if (lower.includes("integrated")) map.igst = i
        else if (lower.includes("central")) map.cgst = i
        else if (lower.includes("state")) map.sgst = i
      } else {
        if (lower.includes("gst no")) map.gstin = i
        else if (lower.includes("ref.no") || lower.includes("invoice no")) map.inv = i
        else if (lower.includes("gst taxable") || lower === "taxable") map.taxable = i
        else if (lower.includes("gst integrated") || lower.includes("integrated")) map.igst = i
        else if (lower.includes("gst central") || lower.includes("central")) map.cgst = i
        else if (lower.includes("gst state") || lower.includes("state")) map.sgst = i
      }
    })

    // Validate critical columns found
    if (map.gstin === undefined || map.inv === undefined) {
      console.warn(`Critical columns missing for ${type}. Headers found:`, headers)
    }

    for (let i = headerRowIndex + 1; i < rows.length; i++) {
      const cols = rows[i]
      if (!cols || cols.length === 0) continue

      const gstin = cols[map.gstin]
      const inv = cols[map.inv]

      // Skip empty rows effectively
      if (!gstin && !inv) continue

      data.push({
        gstin: gstin ? gstin.toString().trim() : "",
        invoice: inv ? inv.toString().trim() : "",
        taxable: parseNumber(cols[map.taxable]),
        igst: parseNumber(cols[map.igst]),
        cgst: parseNumber(cols[map.cgst]),
        sgst: parseNumber(cols[map.sgst]),
      })
    }
    return data
  }

  const handleFileUpload = (e) => {
    const file = e.target.files[0]
    if (!file) return

    if (!libLoaded) {
      setError("Excel processing library is still loading. Please try again in a moment.")
      return
    }

    setIsLoading(true)
    setError("")

    const reader = new FileReader()
    reader.onload = (evt) => {
      try {
        const bstr = evt.target.result
        const wb = window.XLSX.read(bstr, { type: "binary" })

        // 1. Find Sheets
        const sheetNames = wb.SheetNames

        // Flexible matching for sheet names (case insensitive)
        const efactoName = sheetNames.find((n) => n.toLowerCase().includes("efacto"))
        const gstrName = sheetNames.find((n) => n.toLowerCase().includes("2b"))

        if (!efactoName) {
          throw new Error("Could not find a sheet named 'efacto'. Please rename your software data sheet to 'efacto'.")
        }
        if (!gstrName) {
          throw new Error("Could not find a sheet named '2B'. Please rename your GSTR data sheet to '2B'.")
        }

        // 2. Parse Data
        const softSheet = wb.Sheets[efactoName]
        const gstrSheet = wb.Sheets[gstrName]

        const softJson = window.XLSX.utils.sheet_to_json(softSheet, { header: 1 })
        const gstrJson = window.XLSX.utils.sheet_to_json(gstrSheet, { header: 1 })

        const softData = parseSheetData(softJson, "SOFT")
        const gstrData = parseSheetData(gstrJson, "GSTR")

        performReconciliation(softData, gstrData)
      } catch (err) {
        console.error(err)
        setError(err.message || "Failed to process Excel file.")
        setIsLoading(false)
      }
    }
    reader.readAsBinaryString(file)
  }

  const performReconciliation = (softParsed, gstrParsed) => {
    if (gstrParsed.length === 0 && softParsed.length === 0) {
      setError("No valid data found in the sheets. Please check column headers.")
      setIsLoading(false)
      return
    }

    // --- PREPARATION ---
    const gstrMap = {} // For Exact Match (GSTIN + Invoice)
    const gstrByGstin = {} // For Fuzzy Match (GSTIN Only)

    gstrParsed.forEach((row, idx) => {
      row.id = idx // Unique ID for tracking
      row.matched = false // Init matched state

      const cleanInv = cleanString(row.invoice)
      const cleanGst = cleanString(row.gstin)

      const key = `${cleanGst}_${cleanInv}`
      gstrMap[key] = row

      if (!gstrByGstin[cleanGst]) gstrByGstin[cleanGst] = []
      gstrByGstin[cleanGst].push(row)
    })

    const results = []
    const stats = {
      total: 0,
      matched: 0,
      mismatch: 0,
      billMismatch: 0,
      probableMatch: 0,
      missingInGstr: 0,
      missingInSoft: 0,
    }

    // --- PHASE 1: EXACT MATCHING (GSTIN + INVOICE NO) ---
    softParsed.forEach((softRow) => {
      stats.total++
      softRow.matched = false // Init matched state for soft row

      const key = `${cleanString(softRow.gstin)}_${cleanString(softRow.invoice)}`
      const gstrRow = gstrMap[key]

      if (gstrRow && !gstrRow.matched) {
        gstrRow.matched = true
        softRow.matched = true

        // Check Values
        const taxableDiff = Math.abs(softRow.taxable - gstrRow.taxable)
        const taxDiff = Math.abs(
          softRow.igst + softRow.cgst + softRow.sgst - (gstrRow.igst + gstrRow.cgst + gstrRow.sgst),
        )

        if (taxableDiff < 1 && taxDiff < 1) {
          stats.matched++
          results.push({ status: "MATCHED", ...softRow, gstrData: gstrRow, notes: "Perfect Match" })
        } else {
          stats.mismatch++
          results.push({
            status: "MISMATCH",
            ...softRow,
            gstrData: gstrRow,
            notes: `Diff: Taxable ${taxableDiff.toFixed(2)}, Tax ${taxDiff.toFixed(2)}`,
          })
        }
      }
    })

    // --- PHASE 2: BILL MISMATCH (Same Values, Different Invoice) ---
    softParsed.forEach((softRow) => {
      if (softRow.matched) return

      const gstinKey = cleanString(softRow.gstin)
      const potentialMatches = gstrByGstin[gstinKey] || []

      const candidate = potentialMatches.find((gRow) => {
        if (gRow.matched) return false

        const taxableDiff = Math.abs(softRow.taxable - gRow.taxable)
        const taxDiff = Math.abs(softRow.igst + softRow.cgst + softRow.sgst - (gRow.igst + gRow.cgst + gRow.sgst))

        // STRICT value check
        return taxableDiff < 1 && taxDiff < 1
      })

      if (candidate) {
        candidate.matched = true
        softRow.matched = true
        stats.billMismatch++
        results.push({
          status: "BILL_MISMATCH",
          ...softRow,
          gstrData: candidate,
          notes: `Inv No Mismatch (Books: ${softRow.invoice} vs GSTR: ${candidate.invoice})`,
        })
      }
    })

    // --- PHASE 3: PROBABLE MATCH (Similar Invoice + Close Values) ---
    // This handles the "1329" vs "13292" AND "5415" vs "5416" scenario
    softParsed.forEach((softRow) => {
      if (softRow.matched) return

      const gstinKey = cleanString(softRow.gstin)
      const potentialMatches = gstrByGstin[gstinKey] || []

      const candidate = potentialMatches.find((gRow) => {
        if (gRow.matched) return false

        // Check 1: Similar Invoice (Substring or close length)
        const similarInv = isSimilarString(softRow.invoice, gRow.invoice)

        // Check 2: Close Values (Allowing +/- 5.00 tolerance for this fuzzy match)
        const taxableDiff = Math.abs(softRow.taxable - gRow.taxable)
        const taxDiff = Math.abs(softRow.igst + softRow.cgst + softRow.sgst - (gRow.igst + gRow.cgst + gRow.sgst))

        return similarInv && taxableDiff < 5 && taxDiff < 5
      })

      if (candidate) {
        candidate.matched = true
        softRow.matched = true
        stats.probableMatch++
        results.push({
          status: "PROBABLE_MATCH",
          ...softRow,
          gstrData: candidate,
          notes: `Likely Match: Similiar Inv & Close Amt (Diff: ${Math.abs(softRow.taxable - candidate.taxable).toFixed(2)})`,
        })
      } else {
        // If still no match, it's truly missing
        stats.missingInGstr++
        results.push({ status: "MISSING_IN_GSTR", ...softRow, gstrData: null, notes: "Not found in GSTR 2B" })
      }
    })

    // --- PHASE 4: MISSING IN BOOKS ---
    gstrParsed.forEach((gstrRow) => {
      if (!gstrRow.matched) {
        stats.missingInSoft++
        results.push({
          status: "MISSING_IN_BOOKS",
          gstin: gstrRow.gstin,
          invoice: gstrRow.invoice,
          taxable: 0,
          igst: 0,
          cgst: 0,
          sgst: 0,
          gstrData: gstrRow,
          notes: "Not found in Software Data",
        })
      }
    })

    setReconciliationData(results)
    setSummary(stats)
    setIsLoading(false)
    setStep(2)
  }

  const downloadCSV = () => {
    let csvContent = "data:text/csv;charset=utf-8,"
    csvContent +=
      "Status,GSTIN,Invoice No (Books),Invoice No (GSTR),Books Taxable,GSTR Taxable,Books IGST,GSTR IGST,Books CGST,GSTR CGST,Books SGST,GSTR SGST,Difference Note\n"

    reconciliationData.forEach((row) => {
      const rowStr = [
        row.status,
        row.gstin,
        `"${row.invoice}"`,
        row.gstrData ? `"${row.gstrData.invoice}"` : "",
        row.taxable,
        row.gstrData ? row.gstrData.taxable : 0,
        row.igst,
        row.gstrData ? row.gstrData.igst : 0,
        row.cgst,
        row.gstrData ? row.gstrData.cgst : 0,
        row.sgst,
        row.gstrData ? row.gstrData.sgst : 0,
        `"${row.notes}"`,
      ].join(",")
      csvContent += rowStr + "\n"
    })

    const encodedUri = encodeURI(csvContent)
    const link = document.createElement("a")
    link.setAttribute("href", encodedUri)
    link.setAttribute("download", "GSTR_Reconciliation_Report.csv")
    document.body.appendChild(link)
    link.click()
    document.body.removeChild(link)
  }

  return (
    <div className="min-h-screen bg-slate-50 p-4 md:p-8 font-sans">
      <div className="max-w-7xl mx-auto bg-white shadow-xl rounded-xl overflow-hidden border border-slate-200">
        {/* Header */}
        <div className="bg-green-700 p-6 text-white flex justify-between items-center">
          <div>
            <h1 className="text-2xl font-bold flex items-center gap-2">
              <FileSpreadsheet className="h-8 w-8" />
              GSTR vs Books Reconciliation
            </h1>
            <p className="text-green-100 text-sm mt-1">Upload your Excel file containing 'efacto' and '2B' sheets.</p>
          </div>
          {step === 2 && (
            <button
              onClick={() => setStep(1)}
              className="bg-green-800 hover:bg-green-900 text-white px-4 py-2 rounded-lg text-sm font-medium transition flex items-center gap-2"
            >
              <RefreshCw className="h-4 w-4" />
              Start Over
            </button>
          )}
        </div>

        {step === 1 && (
          <div className="p-12 flex flex-col items-center justify-center min-h-[400px]">
            <div className="w-full max-w-xl">
              <label
                className={`
                        flex flex-col items-center justify-center w-full h-64 border-2 border-dashed rounded-lg cursor-pointer 
                        transition-colors duration-300
                        ${isLoading ? "bg-slate-100 border-slate-300 cursor-wait" : "bg-blue-50 border-blue-300 hover:bg-blue-100"}
                    `}
              >
                <div className="flex flex-col items-center justify-center pt-5 pb-6">
                  {isLoading ? (
                    <>
                      <Loader2 className="w-12 h-12 mb-4 text-blue-600 animate-spin" />
                      <p className="mb-2 text-sm text-slate-500 font-semibold">Processing Excel File...</p>
                      <p className="text-xs text-slate-400">Looking for sheets "efacto" and "2B"...</p>
                    </>
                  ) : (
                    <>
                      <Upload className="w-12 h-12 mb-4 text-blue-500" />
                      <p className="mb-2 text-lg text-slate-700 font-bold">Click to upload Excel file</p>
                      <p className="mb-2 text-sm text-slate-500">or drag and drop .xlsx / .xls file</p>
                      <div className="mt-4 text-left text-xs text-slate-400 bg-white p-3 rounded border border-slate-200">
                        <p className="font-semibold mb-1">Requirements:</p>
                        <ul className="list-disc list-inside space-y-1">
                          <li>
                            Sheet 1: <span className="font-mono text-blue-600 font-bold">efacto</span> (Your Software
                            Data)
                          </li>
                          <li>
                            Sheet 2: <span className="font-mono text-blue-600 font-bold">2B</span> (GSTR Data)
                          </li>
                        </ul>
                      </div>
                    </>
                  )}
                </div>
                <input
                  type="file"
                  className="hidden"
                  accept=".xlsx, .xls"
                  onChange={handleFileUpload}
                  disabled={isLoading}
                />
              </label>
            </div>

            {error && (
              <div className="mt-6 p-4 bg-red-50 text-red-700 border border-red-200 rounded-lg flex items-center gap-2">
                <AlertCircle className="h-5 w-5" />
                {error}
              </div>
            )}
            {!libLoaded && !error && <p className="mt-4 text-xs text-slate-400 italic">Initializing Excel engine...</p>}
          </div>
        )}

        {step === 2 && (
          <div className="p-6">
            {/* Summary Cards */}
            <div className="grid grid-cols-2 md:grid-cols-7 gap-2 mb-8">
              <div className="bg-slate-100 p-3 rounded-lg border border-slate-200">
                <div className="text-slate-500 text-[10px] uppercase font-bold">Total Checked</div>
                <div className="text-xl font-bold text-slate-800">{summary.total + summary.missingInSoft}</div>
              </div>
              <div className="bg-green-50 p-3 rounded-lg border border-green-200">
                <div className="text-green-600 text-[10px] uppercase font-bold">Exact Match</div>
                <div className="text-xl font-bold text-green-700">{summary.matched}</div>
              </div>
              <div className="bg-blue-50 p-3 rounded-lg border border-blue-200">
                <div className="text-blue-600 text-[10px] uppercase font-bold">Bill Mismatch</div>
                <div className="text-xl font-bold text-blue-700">{summary.billMismatch}</div>
              </div>
              <div className="bg-yellow-50 p-3 rounded-lg border border-yellow-200">
                <div className="text-yellow-600 text-[10px] uppercase font-bold">Probable Match</div>
                <div className="text-xl font-bold text-yellow-700">{summary.probableMatch}</div>
              </div>
              <div className="bg-red-50 p-3 rounded-lg border border-red-200">
                <div className="text-red-600 text-[10px] uppercase font-bold">Value Mismatch</div>
                <div className="text-xl font-bold text-red-700">{summary.mismatch}</div>
              </div>
              <div className="bg-orange-50 p-3 rounded-lg border border-orange-200">
                <div className="text-orange-600 text-[10px] uppercase font-bold">Not in GSTR</div>
                <div className="text-xl font-bold text-orange-700">{summary.missingInGstr}</div>
              </div>
              <div className="bg-purple-50 p-3 rounded-lg border border-purple-200">
                <div className="text-purple-600 text-[10px] uppercase font-bold">Not in Books</div>
                <div className="text-xl font-bold text-purple-700">{summary.missingInSoft}</div>
              </div>
            </div>

            {/* Actions */}
            <div className="flex justify-between items-center mb-4">
              <div className="flex flex-wrap items-center gap-2 text-xs text-slate-500">
                <span className="flex items-center gap-1">
                  <div className="w-2 h-2 bg-green-500 rounded-full"></div> Exact
                </span>
                <span className="flex items-center gap-1">
                  <div className="w-2 h-2 bg-blue-500 rounded-full"></div> Bill Mismatch
                </span>
                <span className="flex items-center gap-1">
                  <div className="w-2 h-2 bg-yellow-500 rounded-full"></div> Probable (Close Amt & Inv)
                </span>
                <span className="flex items-center gap-1">
                  <div className="w-2 h-2 bg-red-500 rounded-full"></div> Value Mismatch
                </span>
              </div>
              <button
                onClick={downloadCSV}
                className="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-lg text-sm font-medium flex items-center gap-2 shadow"
              >
                <Download className="h-4 w-4" /> Download Report (.csv)
              </button>
            </div>

            {/* Results Table */}
            <div className="overflow-x-auto border rounded-lg max-h-[600px] overflow-y-auto">
              <table className="w-full text-sm text-left relative">
                <thead className="bg-slate-100 text-slate-600 uppercase text-xs font-bold sticky top-0 z-10 shadow-sm">
                  <tr>
                    <th className="px-4 py-3">Status</th>
                    <th className="px-4 py-3">GSTIN</th>
                    <th className="px-4 py-3">Invoice No</th>
                    <th className="px-4 py-3 text-right">Books Taxable</th>
                    <th className="px-4 py-3 text-right">GSTR Taxable</th>
                    <th className="px-4 py-3 text-right">Diff</th>
                    <th className="px-4 py-3 text-right">Total Tax Diff</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {reconciliationData.map((row, idx) => {
                    let statusColor = "bg-gray-100 text-gray-800"
                    let Icon = AlertCircle

                    if (row.status === "MATCHED") {
                      statusColor = "bg-green-100 text-green-800"
                      Icon = CheckCircle
                    } else if (row.status === "BILL_MISMATCH") {
                      statusColor = "bg-blue-100 text-blue-800"
                      Icon = HelpCircle
                    } else if (row.status === "PROBABLE_MATCH") {
                      statusColor = "bg-yellow-100 text-yellow-800"
                      Icon = AlertTriangle
                    } else if (row.status === "MISMATCH") {
                      statusColor = "bg-red-100 text-red-800"
                      Icon = XCircle
                    } else if (row.status === "MISSING_IN_GSTR") {
                      statusColor = "bg-orange-100 text-orange-800"
                      Icon = AlertCircle
                    } else {
                      statusColor = "bg-purple-100 text-purple-800"
                      Icon = AlertCircle
                    }

                    const gstrTaxable = row.gstrData ? row.gstrData.taxable : 0
                    const diffTaxable = row.taxable - gstrTaxable

                    // Calc total tax diff
                    const bookTax = row.igst + row.cgst + row.sgst
                    const gstrTax = row.gstrData ? row.gstrData.igst + row.gstrData.cgst + row.gstrData.sgst : 0
                    const diffTax = bookTax - gstrTax

                    return (
                      <tr key={idx} className="hover:bg-slate-50">
                        <td className="px-4 py-3">
                          <span
                            className={`px-2 py-1 rounded-full text-[10px] font-bold flex items-center gap-1 w-fit ${statusColor}`}
                          >
                            <Icon className="h-3 w-3" />
                            {row.status.replace(/_/g, " ")}
                          </span>
                        </td>
                        <td className="px-4 py-3 font-mono text-slate-600">{row.gstin}</td>
                        <td className="px-4 py-3 font-medium">
                          {row.invoice}
                          {(row.status === "BILL_MISMATCH" || row.status === "PROBABLE_MATCH") && (
                            <div className="text-[10px] text-blue-600 font-mono">GSTR: {row.gstrData?.invoice}</div>
                          )}
                        </td>
                        <td className="px-4 py-3 text-right">{row.taxable.toLocaleString()}</td>
                        <td className="px-4 py-3 text-right text-slate-500">{gstrTaxable.toLocaleString()}</td>
                        <td
                          className={`px-4 py-3 text-right font-bold ${diffTaxable !== 0 ? "text-red-600" : "text-green-600"}`}
                        >
                          {diffTaxable !== 0 ? diffTaxable.toFixed(2) : "-"}
                        </td>
                        <td
                          className={`px-4 py-3 text-right font-bold ${diffTax !== 0 ? "text-red-600" : "text-green-600"}`}
                        >
                          {diffTax !== 0 ? diffTax.toFixed(2) : "-"}
                        </td>
                      </tr>
                    )
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    </div>
  )
}
