"use client"

import type React from "react"

import { useState, useCallback, useRef } from "react"
import { Button } from "@/components/ui/button"
import { Progress } from "@/components/ui/progress"
import { Alert, AlertDescription } from "@/components/ui/alert"
import { FileUp, Download, FileJson, FileSpreadsheet, X } from "lucide-react"
import * as XLSX from "xlsx"

export function JsonToExcel() {
  const [files, setFiles] = useState<File[]>([])
  const [isConverting, setIsConverting] = useState(false)
  const [progress, setProgress] = useState(0)
  const [excelBlob, setExcelBlob] = useState<Blob | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [isDragging, setIsDragging] = useState(false)
  const fileInputRef = useRef<HTMLInputElement>(null)

  // Handle file selection from input
  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    console.log("File input change detected")
    if (e.target.files && e.target.files.length > 0) {
      const fileArray = Array.from(e.target.files).filter((file) => file.name.endsWith(".json"))
      console.log(`Selected ${fileArray.length} JSON files`)
      setFiles(fileArray)
      setExcelBlob(null)
      setError(null)
    }
  }

  // Handle file drop
  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault()
    e.stopPropagation()
    setIsDragging(false)

    console.log("Files dropped")
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const fileArray = Array.from(e.dataTransfer.files).filter((file) => file.name.endsWith(".json"))
      console.log(`Dropped ${fileArray.length} JSON files`)
      setFiles(fileArray)
      setExcelBlob(null)
      setError(null)
    }
  }

  // Handle click on the drop area
  const handleAreaClick = () => {
    console.log("Drop area clicked")
    fileInputRef.current?.click()
  }

  const removeFile = (index: number) => {
    setFiles((prev) => prev.filter((_, i) => i !== index))
    setExcelBlob(null)
  }

  const convertToExcel = useCallback(async () => {
    if (files.length === 0) {
      setError("Por favor, sube al menos un archivo JSON.")
      return
    }

    setIsConverting(true)
    setProgress(0)
    setError(null)

    try {
      // Create a workbook
      const workbook = XLSX.utils.book_new()
      const allData: Record<string, Record<string, string>> = {}

      // Read all JSON files
      for (let i = 0; i < files.length; i++) {
        const file = files[i]
        const locale = file.name.replace(".json", "")

        const text = await file.text()
        const json = JSON.parse(text)

        // Flatten nested JSON if needed
        const flattenedJson = flattenJSON(json)

        // Add to combined data
        for (const key in flattenedJson) {
          if (!allData[key]) {
            allData[key] = {}
          }
          allData[key][locale] = flattenedJson[key]
        }

        setProgress(((i + 1) / files.length) * 100)
      }

      // Convert to worksheet format
      const headers = ["key", ...files.map((f) => f.name.replace(".json", ""))]
      const rows = Object.keys(allData).map((key) => {
        const row: any = { key }
        files.forEach((f) => {
          const locale = f.name.replace(".json", "")
          row[locale] = allData[key][locale] || ""
        })
        return row
      })

      // Create worksheet
      const worksheet = XLSX.utils.json_to_sheet(rows, { header: headers })

      // Add worksheet to workbook
      XLSX.utils.book_append_sheet(workbook, worksheet, "Localization")

      // Generate Excel file
      const excelBuffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" })
      const blob = new Blob([excelBuffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      })

      setExcelBlob(blob)
      setProgress(100)
    } catch (err) {
      console.error(err)
      setError("Error al convertir los archivos. Verifica que sean archivos JSON válidos.")
    } finally {
      setIsConverting(false)
    }
  }, [files])

  const downloadExcel = () => {
    if (excelBlob) {
      const url = URL.createObjectURL(excelBlob)
      const a = document.createElement("a")
      a.href = url
      a.download = "localization.xlsx"
      document.body.appendChild(a)
      a.click()
      document.body.removeChild(a)
      URL.revokeObjectURL(url)
    }
  }

  // Function to flatten nested JSON objects
  const flattenJSON = (obj: any, prefix = ""): Record<string, string> => {
    const result: Record<string, string> = {}

    for (const key in obj) {
      const newKey = prefix ? `${prefix}.${key}` : key

      if (typeof obj[key] === "object" && obj[key] !== null && !Array.isArray(obj[key])) {
        Object.assign(result, flattenJSON(obj[key], newKey))
      } else {
        result[newKey] = String(obj[key])
      }
    }

    return result
  }

  return (
    <div className="space-y-6">
      {/* Hidden file input */}
      <input
        ref={fileInputRef}
        type="file"
        multiple
        accept=".json"
        onChange={handleFileChange}
        className="hidden"
        id="json-file-upload"
      />

      {/* Drag and drop area */}
      <div
        className={`flex flex-col items-center justify-center border-2 border-dashed rounded-lg p-12 text-center cursor-pointer transition-colors ${
          isDragging ? "border-primary bg-primary/5" : "border-gray-300 hover:border-gray-400"
        }`}
        onClick={handleAreaClick}
        onDragOver={(e) => {
          e.preventDefault()
          e.stopPropagation()
          setIsDragging(true)
        }}
        onDragLeave={(e) => {
          e.preventDefault()
          e.stopPropagation()
          setIsDragging(false)
        }}
        onDrop={handleDrop}
      >
        <FileJson className="h-10 w-10 text-gray-400 mb-4" />
        <div className="space-y-2">
          <p className="text-sm text-gray-500">Arrastra y suelta archivos JSON o haz clic para seleccionarlos</p>
          <p className="text-xs text-gray-400">Cada archivo JSON debe representar un idioma</p>
        </div>
        <Button variant="outline" className="mt-4">
          <FileUp className="mr-2 h-4 w-4" />
          Seleccionar archivos
        </Button>
      </div>

      {files.length > 0 && (
        <div className="space-y-4">
          <h3 className="font-medium">Archivos seleccionados ({files.length})</h3>
          <ul className="space-y-2">
            {files.map((file, index) => (
              <li key={index} className="flex items-center justify-between p-2 bg-gray-50 rounded">
                <span className="text-sm truncate max-w-[80%]">{file.name}</span>
                <Button variant="ghost" size="icon" onClick={() => removeFile(index)} className="h-8 w-8">
                  <X className="h-4 w-4" />
                </Button>
              </li>
            ))}
          </ul>

          {isConverting ? (
            <div className="space-y-2">
              <p className="text-sm">Convirtiendo...</p>
              <Progress value={progress} />
            </div>
          ) : (
            <div className="flex space-x-2">
              <Button onClick={convertToExcel} disabled={files.length === 0}>
                Convertir a Excel
              </Button>
              {excelBlob && (
                <Button variant="outline" onClick={downloadExcel}>
                  <Download className="mr-2 h-4 w-4" />
                  Descargar Excel
                </Button>
              )}
            </div>
          )}
        </div>
      )}

      {error && (
        <Alert variant="destructive">
          <AlertDescription>{error}</AlertDescription>
        </Alert>
      )}

      {excelBlob && (
        <Alert>
          <FileSpreadsheet className="h-4 w-4 mr-2" />
          <AlertDescription>
            ¡Conversión completada! Haz clic en "Descargar Excel" para guardar el archivo.
          </AlertDescription>
        </Alert>
      )}
    </div>
  )
}
