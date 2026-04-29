"use client";

import React, { useState } from "react";
import {
  Upload,
  FileCheck,
  Loader2,
  Download,
  AlertCircle,
  FileText,
  Sparkles,
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
    if (downloadUrl) window.URL.revokeObjectURL(downloadUrl);
  };

  const parseBlobError = async function (blob) {
    try {
      const text = await blob.text();
      const data = JSON.parse(text);
      return data && data.detail ? data.detail : "Unknown server error";
    } catch {
      return "The server returned an unreadable error response.";
    }
  };

  const handleUpload = async function () {
    if (!guidelines || !unformattedDocument) {
      setError("Please upload both the style guide and the document to format.");
      return;
    }

    setLoading(true);
    setError(null);
    revokeDownloadUrl();
    setDownloadUrl(null);

    const formData = new FormData();
    formData.append("guidelines", guidelines);
    formData.append("unformatted_document", unformattedDocument);
    if (referenceDocument) formData.append("reference_document", referenceDocument);

    try {
      const response = await axios.post(API_BASE_URL + "/v1/format", formData, {
        responseType: "blob",
        timeout: 600000,
      });

      const blob = new Blob([response.data], {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });
      setDownloadUrl(window.URL.createObjectURL(blob));
    } catch (err) {
      if (axios.isAxiosError(err)) {
        if (err.response?.data instanceof Blob) {
          const detail = await parseBlobError(err.response.data);
          setError("Server error: " + detail);
        } else if (err.response) {
          setError("Server error: Request failed.");
        } else if (err.request) {
          setError("Could not reach the server. Make sure the API is running.");
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

  const renderUploadBox = function (label, file, onChange, accent) {
    const accents = {
      indigo: {
        active: "border-indigo-500 bg-indigo-500/10 text-indigo-400",
        icon: "text-indigo-400",
      },
      violet: {
        active: "border-violet-500 bg-violet-500/10 text-violet-400",
        icon: "text-violet-400",
      },
    };
    const a = accents[accent];

    return (
      <div
        className={
          "p-5 border-2 border-dashed rounded-2xl transition-all duration-200 " +
          (file ? a.active : "border-zinc-800 bg-zinc-900/40 hover:border-zinc-600")
        }
      >
        <label className="flex flex-col items-center gap-2 cursor-pointer">
          {file ? (
            <FileCheck size={22} className={a.icon} />
          ) : (
            <Upload size={22} className="text-zinc-600" />
          )}
          <span className="text-xs text-center leading-relaxed text-zinc-400">
            {file ? file.name : label}
          </span>
          <input
            type="file"
            hidden
            accept=".docx"
            onChange={function (e) {
              onChange(e.target.files?.[0] ?? null);
            }}
          />
        </label>
      </div>
    );
  };

  return (
    <Card className="w-full max-w-xl bg-zinc-950/80 border border-zinc-800/60 text-white shadow-2xl rounded-3xl backdrop-blur-sm">
      <CardHeader className="pb-2">
        <CardTitle className="text-lg font-semibold text-center text-zinc-200 tracking-tight">
          Format a Document
        </CardTitle>
        <p className="text-center text-zinc-600 text-xs mt-1">
          Provide a style guide and a document — AI handles the rest.
        </p>
      </CardHeader>

      <CardContent className="space-y-4 pt-2">
        {/* Main uploads */}
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          {renderUploadBox("Style Guide (.docx)", guidelines, setGuidelines, "indigo")}
          {renderUploadBox("Document to Format (.docx)", unformattedDocument, setUnformattedDocument, "violet")}
        </div>

        {/* Optional reference */}
        <div className="p-4 border border-zinc-800 rounded-2xl bg-zinc-900/40 hover:border-zinc-700 transition-colors">
          <label className="flex items-center gap-3 cursor-pointer">
            <FileText size={18} className={referenceDocument ? "text-amber-400" : "text-zinc-600"} />
            <div className="flex-1 min-w-0">
              <div className="text-xs font-medium text-zinc-300 truncate">
                {referenceDocument ? referenceDocument.name : "Reference Example (optional)"}
              </div>
              <div className="text-xs text-zinc-600 mt-0.5">
                Helps the AI match a specific output style.
              </div>
            </div>
            <input
              type="file"
              hidden
              accept=".docx"
              onChange={function (e) {
                setReferenceDocument(e.target.files?.[0] ?? null);
              }}
            />
          </label>
        </div>

        {/* Error */}
        {error && (
          <div className="bg-red-500/10 border border-red-500/40 text-red-400 p-3 rounded-xl flex items-start gap-2 text-xs">
            <AlertCircle size={14} className="mt-0.5 shrink-0" />
            <span>{error}</span>
          </div>
        )}

        {/* Action */}
        {!downloadUrl ? (
          <Button
            onClick={handleUpload}
            disabled={!canSubmit}
            className="w-full py-5 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-white font-semibold text-sm transition-all duration-200 disabled:opacity-30 disabled:cursor-not-allowed"
          >
            {loading ? (
              <>
                <Loader2 size={15} className="animate-spin mr-2" />
                Formatting with AI…
              </>
            ) : (
              <>
                <Sparkles size={15} className="mr-2" />
                Run Formatter
              </>
            )}
          </Button>
        ) : (
          <div className="space-y-2">
            <a href={downloadUrl} download="formatted_document.docx">
              <Button className="w-full rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white font-semibold text-sm py-5">
                <Download size={15} className="mr-2" />
                Download Formatted Doc
              </Button>
            </a>
            <Button
              variant="ghost"
              onClick={reset}
              className="w-full text-zinc-500 hover:text-zinc-300 text-xs"
            >
              Format another document
            </Button>
          </div>
        )}
      </CardContent>
    </Card>
  );
}