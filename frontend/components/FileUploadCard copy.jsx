"use client";

import React, { useState } from "react";
import {
  Upload,
  FileCheck,
  Loader2,
  Download,
  AlertCircle,
  FileText,
} from "lucide-react";
import axios from "axios";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";

const API_BASE_URL = process.env.NEXT_PUBLIC_API_BASE_URL || "http://localhost:8000";

export default function FileUploadCard() {
  const [guidelines, setGuidelines] = useState(null);
  const [unformattedDocument, setUnformattedDocument] = useState(null);
  const [referenceDocument, setReferenceDocument] = useState(null);
  const [loading, setLoading] = useState(false);
  const [downloadUrl, setDownloadUrl] = useState(null);
  const [error, setError] = useState(null);

  const canSubmit = !!guidelines && !!unformattedDocument && !loading;

  const revokeDownloadUrl = function () {
    if (downloadUrl) {
      window.URL.revokeObjectURL(downloadUrl);
    }
  };

  const parseBlobError = async function (blob) {
    try {
      const text = await blob.text();
      const data = JSON.parse(text);
      return data && data.detail ? data.detail : "Unknown server error";
    } catch (parseError) {
      return "The server returned an unreadable error response.";
    }
  };

  const handleUpload = async function () {
    if (!guidelines || !unformattedDocument) {
      setError("Please upload both the guidelines document and the unformatted document.");
      return;
    }

    setLoading(true);
    setError(null);
    revokeDownloadUrl();
    setDownloadUrl(null);

    const formData = new FormData();
    formData.append("guidelines", guidelines);
    formData.append("unformatted_document", unformattedDocument);

    if (referenceDocument) {
      formData.append("reference_document", referenceDocument);
    }

    try {
      const response = await axios.post(API_BASE_URL + "/v1/format", formData, {
        responseType: "blob",
        timeout: 600000,
      });

      const blob = new Blob([response.data], {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });
      const url = window.URL.createObjectURL(blob);
      setDownloadUrl(url);
    } catch (err) {
      if (axios.isAxiosError(err)) {
        if (err.response && err.response.data && err.response.data instanceof Blob) {
          const detail = await parseBlobError(err.response.data);
          setError("Server error: " + detail);
        } else if (err.response) {
          setError("Server error: Request failed.");
        } else if (err.request) {
          setError("Backend is not responding. Check whether the API server is running.");
        } else {
          setError(err.message || "Unexpected error occurred.");
        }
      } else {
        setError("Unexpected error occurred.");
      }
    } finally {
      setLoading(false);
    }
  };

  const reset = function () {
    revokeDownloadUrl();
    setGuidelines(null);
    setUnformattedDocument(null);
    setReferenceDocument(null);
    setDownloadUrl(null);
    setError(null);
  };

  const renderUploadBox = function (label, file, onChange, accentClasses) {
    return (
      <div
        className={"p-6 border-2 border-dashed rounded-xl transition-colors " +
          (file ? accentClasses : "border-zinc-700 bg-zinc-900/40")}
      >
        <label className="flex flex-col items-center cursor-pointer">
          {file ? <FileCheck className="text-current" /> : <Upload className="text-zinc-500" />}
          <span className="text-sm mt-2 text-center">{file ? file.name : label}</span>
          <input
            type="file"
            hidden
            accept=".docx"
            onChange={function (e) {
              const selectedFile = e.target.files && e.target.files[0] ? e.target.files[0] : null;
              onChange(selectedFile);
            }}
          />
        </label>
      </div>
    );
  };

  return (
    <Card className="w-full max-w-2xl bg-zinc-950 border-zinc-800 text-white shadow-2xl">
      <CardHeader>
        <CardTitle className="text-2xl font-bold text-center">
          Legal Document Formatter
        </CardTitle>
      </CardHeader>

      <CardContent className="space-y-6">
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          {renderUploadBox(
            "Upload Guidelines (.docx)",
            guidelines,
            setGuidelines,
            "border-emerald-500 bg-emerald-500/10 text-emerald-400"
          )}

          {renderUploadBox(
            "Upload Unformatted Document (.docx)",
            unformattedDocument,
            setUnformattedDocument,
            "border-indigo-500 bg-indigo-500/10 text-indigo-400"
          )}
        </div>

        <div className="p-4 border border-zinc-800 rounded-xl bg-zinc-900/60">
          <label className="flex items-center gap-3 cursor-pointer">
            <FileText className={referenceDocument ? "text-amber-400" : "text-zinc-500"} />
            <div className="flex-1">
              <div className="text-sm font-medium">
                {referenceDocument
                  ? referenceDocument.name
                  : "Optional: Upload Reference Formatted Document (.docx)"}
              </div>
              <div className="text-xs text-zinc-400">
                Use this when you want the output to more closely match an example document.
              </div>
            </div>
            <input
              type="file"
              hidden
              accept=".docx"
              onChange={function (e) {
                const selectedFile = e.target.files && e.target.files[0] ? e.target.files[0] : null;
                setReferenceDocument(selectedFile);
              }}
            />
          </label>
        </div>

        {error && (
          <div className="bg-red-500/10 border border-red-500 text-red-400 p-3 rounded flex items-center gap-2">
            <AlertCircle size={16} />
            <span>{error}</span>
          </div>
        )}

        {!downloadUrl ? (
          <Button
            onClick={handleUpload}
            disabled={!canSubmit}
            className="w-full py-6 bg-white text-black hover:bg-zinc-200"
          >
            {loading ? (
              <>
                <Loader2 className="animate-spin mr-2" />
                Processing document...
              </>
            ) : (
              "Start Formatting"
            )}
          </Button>
        ) : (
          <div className="space-y-3">
            <a href={downloadUrl} download="formatted_document.docx">
              <Button className="w-full bg-emerald-600 hover:bg-emerald-500">
                <Download className="mr-2" />
                Download File
              </Button>
            </a>

            <Button variant="ghost" onClick={reset}>
              Format Another
            </Button>
          </div>
        )}
      </CardContent>
    </Card>
  );
}
