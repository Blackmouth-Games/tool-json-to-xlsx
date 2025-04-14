"use client"
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs"
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card"
import { JsonToExcel } from "@/components/json-to-excel"
import { ExcelToJson } from "@/components/excel-to-json"

export default function Home() {
  return (
    <main className="container mx-auto py-10 px-4">
      <h1 className="text-3xl font-bold text-center mb-8">Localization File Converter</h1>

      <Tabs defaultValue="json-to-excel" className="max-w-4xl mx-auto">
        <TabsList className="grid w-full grid-cols-2">
          <TabsTrigger value="json-to-excel">JSON a Excel</TabsTrigger>
          <TabsTrigger value="excel-to-json">Excel a JSON</TabsTrigger>
        </TabsList>

        <TabsContent value="json-to-excel">
          <Card>
            <CardHeader>
              <CardTitle>Convertir archivos JSON a Excel</CardTitle>
              <CardDescription>
                Sube múltiples archivos JSON de localización y conviértelos en un único archivo Excel.
              </CardDescription>
            </CardHeader>
            <CardContent>
              <JsonToExcel />
            </CardContent>
          </Card>
        </TabsContent>

        <TabsContent value="excel-to-json">
          <Card>
            <CardHeader>
              <CardTitle>Convertir Excel a archivos JSON</CardTitle>
              <CardDescription>
                Sube un archivo Excel con columnas para cada idioma y conviértelo en múltiples archivos JSON.
              </CardDescription>
            </CardHeader>
            <CardContent>
              <ExcelToJson />
            </CardContent>
          </Card>
        </TabsContent>
      </Tabs>
    </main>
  )
}
