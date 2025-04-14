"use client"

import type React from "react"

import { useState, useCallback, useRef } from "react"
import { Button } from "@/components/ui/button"
import { Progress } from "@/components/ui/progress"
import { Alert, AlertDescription } from "@/components/ui/alert"
import { FileUp, Download, FileJson, FileSpreadsheet, X } from "lucide-react"
import * as XLSX from "xlsx"
import JSZip from "jszip"

export function ExcelToJson() {
  const [file, setFile] = useState<File | null>(null)
  const [isConverting, setIsConverting] = useState(false)
  const [progress, setProgress] = useState(0)
  const [jsonZipBlob, setJsonZipBlob] = useState<Blob | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [isDragging, setIsDragging] = useState(false)
  const [previewData, setPreviewData] = useState<{
    headers: string[]
    rowCount: number
  } | null>(null)
  const fileInputRef = useRef<HTMLInputElement>(null)

  const handleExcelPreview = async (selectedFile: File) => {
    try {
      console.log("Previewing Excel file:", selectedFile.name)
      // Preview the Excel file
      const data = await selectedFile.arrayBuffer()
      const workbook = XLSX.read(data)
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]]
      const jsonData = XLSX.utils.sheet_to_json(firstSheet)

      if (jsonData.length === 0) {
        setError("El archivo Excel está vacío")
        setFile(null)
        return
      }

      // Get headers (first row)
      const headers = Object.keys(jsonData[0])

      if (!headers.includes("key")) {
        setError("El archivo Excel debe contener una columna 'key'")
        setFile(null)
        return
      }

      setPreviewData({
        headers,
        rowCount: jsonData.length,
      })
      console.log("Excel preview successful:", headers)
    } catch (err) {
      console.error("Excel preview error:", err)
      setError("Error al leer el archivo Excel. Verifica que sea un archivo válido.")
      setFile(null)
    }
  }

  // Handle file selection from input
  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    console.log("File input change detected")
    if (e.target.files && e.target.files[0]) {
      const selectedFile = e.target.files[0]
      console.log("Selected file:", selectedFile.name)

      if (!selectedFile.name.endsWith(".xlsx") && !selectedFile.name.endsWith(".xls")) {
        setError("Por favor, sube un archivo Excel válido (.xlsx o .xls)")
        return
      }

      setFile(selectedFile)
      setJsonZipBlob(null)
      setError(null)

      // Preview the Excel file
      await handleExcelPreview(selectedFile)
    }
  }

  // Handle file drop
  const handleDrop = async (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault()
    e.stopPropagation()
    setIsDragging(false)

    console.log("Files dropped")
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const selectedFile = e.dataTransfer.files[0]
      console.log("Dropped file:", selectedFile.name)

      if (!selectedFile.name.endsWith(".xlsx") && !selectedFile.name.endsWith(".xls")) {
        setError("Por favor, sube un archivo Excel válido (.xlsx o .xls)")
        return
      }

      setFile(selectedFile)
      setJsonZipBlob(null)
      setError(null)

      // Preview the Excel file
      await handleExcelPreview(selectedFile)
    }
  }

  // Handle click on the drop area
  const handleAreaClick = () => {
    console.log("Drop area clicked")
    fileInputRef.current?.click()
  }

  const removeFile = () => {
    setFile(null)
    setPreviewData(null)
    setJsonZipBlob(null)
  }

  const convertToJson = useCallback(async () => {
    if (!file) {
      setError("Por favor, sube un archivo Excel.")
      return
    }

    setIsConverting(true)
    setProgress(0)
    setError(null)

    try {
      // Read Excel file
      const data = await file.arrayBuffer()
      const workbook = XLSX.read(data)
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]]
      const jsonData = XLSX.utils.sheet_to_json(firstSheet)

      // Check if there's a key column
      if (!jsonData[0] || !("key" in jsonData[0])) {
        throw new Error("El archivo Excel debe contener una columna 'key'")
      }

      // Get all locales (columns except 'key')
      const locales = Object.keys(jsonData[0]).filter((key) => key !== "key")

      if (locales.length === 0) {
        throw new Error("El archivo Excel debe contener al menos una columna de idioma además de 'key'")
      }

      // Create a JSON file for each locale
      const zip = new JSZip()

      for (let i = 0; i < locales.length; i++) {
        const locale = locales[i]
        const localeData: Record<string, any> = {}

        // Process each row
        jsonData.forEach((row: any) => {
          if (row.key && row[locale] !== undefined) {
            // Handle nested keys (e.g., "common.button.submit")
            setNestedValue(localeData, row.key, row[locale])
          }
        })

        // Add to zip
        zip.file(`${locale}.json`, JSON.stringify(localeData, null, 2))

        setProgress(((i + 1) / locales.length) * 100)
      }

      // Generate zip file
      const zipBlob = await zip.generateAsync({ type: "blob" })
      setJsonZipBlob(zipBlob)
      setProgress(100)
    } catch (err: any) {
      console.error(err)
      setError(err.message || "Error al convertir el archivo Excel a JSON.")
    } finally {
      setIsConverting(false)
    }
  }, [file])

  const downloadJsonZip = () => {
    if (jsonZipBlob) {
      const url = URL.createObjectURL(jsonZipBlob)
      const a = document.createElement("a")
      a.href = url
      a.download = "localization_json_files.zip"
      document.body.appendChild(a)
      a.click()
      document.body.removeChild(a)
      URL.revokeObjectURL(url)
    }
  }

  // Function to set nested values in an object based on a dot-notation key
  const setNestedValue = (obj: Record<string, any>, path: string, value: any) => {
    const keys = path.split(".")
    let current = obj

    for (let i = 0; i < keys.length - 1; i++) {
      const key = keys[i]
      if (!current[key]) {
        current[key] = {}
      }
      current = current[key]
    }

    current[keys[keys.length - 1]] = value
  }

  return (
    <div className="space-y-6">
      {/* Hidden file input */}
      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx,.xls"
        onChange={handleFileChange}
        className="hidden"
        id="excel-file-upload"
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
        <FileSpreadsheet className="h-10 w-10 text-gray-400 mb-4" />
        <div className="space-y-2">
          <p className="text-sm text-gray-500">Arrastra y suelta un archivo Excel o haz clic para seleccionarlo</p>
          <p className="text-xs text-gray-400">El archivo debe tener una columna "key" y columnas para cada idioma</p>
        </div>
        <Button variant="outline" className="mt-4">
          <FileUp className="mr-2 h-4 w-4" />
          Seleccionar archivo
        </Button>
      </div>

      {file && (
        <div className="space-y-4">
          <div className="flex items-center justify-between p-2 bg-gray-50 rounded">
            <span className="text-sm truncate max-w-[80%]">{file.name}</span>
            <Button variant="ghost" size="icon" onClick={removeFile} className="h-8 w-8">
              <X className="h-4 w-4" />
            </Button>
          </div>

          {previewData && (
            <div className="bg-gray-50 p-3 rounded text-sm">
              <p>
                <strong>Columnas detectadas:</strong> {previewData.headers.join(", ")}
              </p>
              <p>
                <strong>Número de filas:</strong> {previewData.rowCount}
              </p>
              <p>
                <strong>Idiomas:</strong> {previewData.headers.filter((h) => h !== "key").join(", ")}
              </p>
            </div>
          )}

          {isConverting ? (
            <div className="space-y-2">
              <p className="text-sm">Convirtiendo...</p>
              <Progress value={progress} />
            </div>
          ) : (
            <div className="flex space-x-2">
              <Button onClick={convertToJson} disabled={!file}>
                Convertir a JSON
              </Button>
              {jsonZipBlob && (
                <Button variant="outline" onClick={downloadJsonZip}>
                  <Download className="mr-2 h-4 w-4" />
                  Descargar archivos JSON
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

      {jsonZipBlob && (
        <Alert>
          <FileJson className="h-4 w-4 mr-2" />
          <AlertDescription>
            ¡Conversión completada! Haz clic en "Descargar archivos JSON" para guardar el archivo ZIP.
          </AlertDescription>
        </Alert>
      )}
    </div>
  )
}
