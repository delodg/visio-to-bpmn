"use client";

import React, { useEffect, useRef, useState } from "react";
import BpmnJS from "bpmn-js";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import { XMLParser, XMLBuilder } from "fast-xml-parser";
import JSZip from "jszip";
import { VisioToBpmnConverterClass } from "./visio-to-bpmn-converter-class";
import BpmnModdle from "bpmn-moddle";

export function VisioToBpmnConverter() {
  const [file, setFile] = useState<File | null>(null);
  const [convertedXml, setConvertedXml] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const bpmnViewerRef = useRef<HTMLDivElement>(null);
  const bpmnViewer = useRef<BpmnJS | null>(null);

  useEffect(() => {
    if (bpmnViewerRef.current && !bpmnViewer.current) {
      bpmnViewer.current = new BpmnJS({ container: bpmnViewerRef.current });
    }

    return () => {
      if (bpmnViewer.current) {
        bpmnViewer.current.destroy();
      }
    };
  }, []);

  useEffect(() => {
    if (convertedXml && bpmnViewer.current) {
      const moddle = new BpmnModdle();

      moddle
        .fromXML(convertedXml)
        .then(({ warnings }) => {
          if (warnings.length) {
            console.warn("Warnings during XML import:", warnings);
          }
          return bpmnViewer.current.importXML(convertedXml);
        })
        .then(() => {
          if (bpmnViewer.current) {
            bpmnViewer.current.get("canvas").zoom("fit-viewport");
          }
        })
        .catch((err: Error) => {
          console.error("Error rendering BPMN diagram", err);
          setError(`Error rendering BPMN diagram: ${err.message}`);
        });
    }
  }, [convertedXml]);

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    if (event.target.files) {
      setFile(event.target.files[0]);
      setConvertedXml(null);
      setError(null);
    }
  };

  const handleConvert = async () => {
    if (!file) {
      setError("Please select a Visio file first.");
      return;
    }

    try {
      const fileContent = await readFileAsArrayBuffer(file);
      console.log("File content loaded as ArrayBuffer");

      const converter = new VisioToBpmnConverterClass();
      const bpmnXml = await converter.convert(fileContent);
      console.log("Converted BPMN XML (first 500 chars):", bpmnXml.substring(0, 500));

      setConvertedXml(bpmnXml);
      setError(null);
    } catch (err) {
      console.error("Full error:", err);
      setError(`Error converting file: ${err.message}`);
    }
  };

  const handleDownload = () => {
    if (convertedXml) {
      const blob = new Blob([convertedXml], { type: "text/xml" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "converted_bpmn.xml";
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }
  };

  return (
    <Card className="w-full max-w-4xl mx-auto">
      <CardHeader>
        <CardTitle>Visio to BPMN Converter</CardTitle>
        <CardDescription>Upload a Visio XML file (.vdx or .vsdx) to convert it to BPMN format</CardDescription>
      </CardHeader>
      <CardContent>
        <div className="space-y-4">
          <Input type="file" accept=".xml,.vdx,.vsdx" onChange={handleFileChange} />
          <Button onClick={handleConvert} className="w-full">
            Convert
          </Button>
          {convertedXml && (
            <Button onClick={handleDownload} className="w-full">
              Download BPMN XML
            </Button>
          )}
          {error && (
            <Alert variant="destructive">
              <AlertTitle>Error</AlertTitle>
              <AlertDescription>{error}</AlertDescription>
            </Alert>
          )}
          <div
            ref={bpmnViewerRef}
            style={{
              height: "400px",
              border: "1px solid #ccc",
              display: convertedXml ? "block" : "none",
            }}
          ></div>
        </div>
      </CardContent>
    </Card>
  );
}

function readFileAsArrayBuffer(file: File): Promise<ArrayBuffer> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = event => resolve(event.target?.result as ArrayBuffer);
    reader.onerror = error => reject(error);
    reader.readAsArrayBuffer(file);
  });
}
