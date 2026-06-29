"use client";

import { useState, useRef } from "react";

enum LogLevel {
  "LIMITED" = 0,
  "TRANSCRIPTION" = 1,
  "FULL" = 2,
}

type ProcessStatus =
  | "tos"
  | "idle"
  | "uploading"
  | "processing"
  | "success"
  | "error";

interface ProcessedFile {
  name: string;
  size: number;
  mimetype: string;
  data: string; // base64 string
}

export default function Home() {
  const [status, setStatus] = useState<ProcessStatus>("tos");
  const [logLevel, setLogLevel] = useState<LogLevel>(LogLevel.TRANSCRIPTION);
  const [file, setFile] = useState<File | null>(null);
  const [processedFile, setProcessedFile] = useState<ProcessedFile | null>(
    null,
  );
  const [errorMsg, setErrorMsg] = useState<string>("");
  const [isDragOver, setIsDragOver] = useState(false);

  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(true);
  };

  const handleDragLeave = () => {
    setIsDragOver(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);

    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const droppedFile = e.dataTransfer.files[0];
      validateAndSetFile(droppedFile);
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      validateAndSetFile(e.target.files[0]);
    }
  };

  const validateAndSetFile = (selectedFile: File) => {
    const isDocx = selectedFile.name.endsWith(".docx");
    const isPptx = selectedFile.name.endsWith(".pptx");
    if (!isDocx && !isPptx) {
      setErrorMsg("Invalid file type. Please upload a .docx Word document or a .pptx PowerPoint presentation.");
      setStatus("error");
      setFile(null);
      return;
    }
    setFile(selectedFile);
    setErrorMsg("");
    setStatus("idle");
  };

  const triggerFileSelect = () => {
    fileInputRef.current?.click();
  };

  const startProcessing = async () => {
    if (!file) return;

    try {
      setStatus("uploading");
      setErrorMsg("");

      // 1. Upload file (POST /api/upload)
      const uploadForm = new FormData();
      uploadForm.append("uploadedFile", file);
      uploadForm.append("logLevel", logLevel.toString());

      const uploadRes = await fetch("/api/upload", {
        method: "POST",
        body: uploadForm,
      });

      if (!uploadRes.ok) {
        const errJson = await uploadRes.json().catch(() => ({}));
        throw new Error(
          errJson.error || `Upload failed with status ${uploadRes.status}`,
        );
      }

      const uploadData = await uploadRes.json();
      const jobId = uploadData.jobId;

      if (!jobId) {
        throw new Error("API did not return a valid Job ID.");
      }

      // 2. Subscribe and wait for completion (GET /api/subscribe?jobId=...)
      setStatus("processing");

      const fileType = file.name.endsWith(".pptx") ? "pptx" : "docx";
      const subscribeRes = await fetch(
        `/api/subscribe?jobId=${encodeURIComponent(jobId)}&fileType=${fileType}`,
      );

      if (!subscribeRes.ok) {
        const errJson = await subscribeRes.json().catch(() => ({}));
        throw new Error(
          errJson.error ||
            `Processing failed with status ${subscribeRes.status}`,
        );
      }

      const subscribeData = await subscribeRes.json();

      if (!subscribeData.data) {
        throw new Error("Processing completed but no file data was returned.");
      }

      const isPptx = file.name.endsWith(".pptx");
      setProcessedFile({
        name:
          subscribeData.name ||
          (isPptx
            ? file.name.replace(".pptx", "_processed.pptx")
            : file.name.replace(".docx", "_processed.docx")),
        size: subscribeData.size || file.size,
        mimetype:
          subscribeData.mimetype ||
          (isPptx
            ? "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            : "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
        data: subscribeData.data,
      });

      setStatus("success");
    } catch (err: unknown) {
      console.error(err);
      const errMsg = err instanceof Error ? err.message : "An unexpected error occurred while processing the document.";
      setErrorMsg(errMsg);
      setStatus("error");
    }
  };

  const downloadProcessedFile = () => {
    if (!processedFile) return;

    try {
      // Decode base64
      const base64Clean = processedFile.data.includes("base64,")
        ? processedFile.data.split("base64,")[1]
        : processedFile.data;

      const byteCharacters = atob(base64Clean);
      const byteNumbers = new Array(byteCharacters.length);
      for (let i = 0; i < byteCharacters.length; i++) {
        byteNumbers[i] = byteCharacters.charCodeAt(i);
      }
      const byteArray = new Uint8Array(byteNumbers);
      const blob = new Blob([byteArray], { type: processedFile.mimetype });

      // Create download link
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = processedFile.name;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error("Download failed:", err);
      alert("Failed to decode and download the file. Please try again.");
    }
  };

  const resetState = () => {
    setStatus("idle");
    setFile(null);
    setProcessedFile(null);
    setErrorMsg("");
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
  };

  const formatBytes = (bytes: number, decimals = 2) => {
    if (bytes === 0) return "0 Bytes";
    const k = 1024;
    const dm = decimals < 0 ? 0 : decimals;
    const sizes = ["Bytes", "KB", "MB", "GB"];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + " " + sizes[i];
  };

  return (
    <div className="relative min-h-screen bg-zinc-950 bg-grid-pattern text-zinc-100 flex flex-col items-center justify-between pb-12 font-sans selection:bg-violet-500 selection:text-white">
      {/* Visual Glowing Orb */}
      <div className="absolute top-0 inset-x-0 h-[500px] radial-glow pointer-events-none" />

      {/* Header */}
      <header className="relative w-full max-w-5xl px-6 pt-8 pb-4 flex items-center justify-between border-b border-zinc-900 z-10">
        <div className="flex items-center gap-3">
          <div className="h-10 w-10 rounded-xl bg-gradient-to-tr from-violet-600 to-indigo-600 flex items-center justify-center shadow-lg shadow-violet-500/20">
            <svg
              className="h-5 w-5 text-white"
              fill="none"
              stroke="currentColor"
              viewBox="0 0 24 24"
              xmlns="http://www.w3.org/2000/svg"
            >
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                strokeWidth={2}
                d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"
              />
            </svg>
          </div>
          <div>
            <span className="font-bold text-lg tracking-tight bg-gradient-to-r from-white via-zinc-200 to-zinc-400 bg-clip-text text-transparent">
              EduAlly
            </span>
            <span className="text-xs block text-violet-400 font-semibold tracking-wider uppercase">
              Accessibility Engine
            </span>
          </div>
        </div>
        <div className="flex items-center gap-4">
          <span className="text-xs text-zinc-500 bg-zinc-900 px-3 py-1 rounded-full border border-zinc-800">
            Beta
          </span>
        </div>
      </header>

      {/* Main Container */}
      <main className="relative w-full max-w-2xl px-6 mt-16 flex-1 flex flex-col justify-center z-10">
        {/* State 0: TOS query */}
        {status === "tos" && (
          <div className="space-y-8 animate-fade-in">
            <div className="text-center space-y-3">
              <h1 className="text-4xl font-extrabold tracking-tight text-white sm:text-5xl">
                Data Collection Preference
              </h1>
              <p className="text-zinc-400 max-w-lg mx-auto text-base">
                Please select which level of data collection you are most
                comfortable with. Your choice helps us respect your data privacy
                while contributing to our research and improving our services.
                Any data you choose to submit is securely stored and is not used
                for any purpose other than accessibility analysis.
              </p>
            </div>

            <div className="grid grid-cols-1 gap-4 sm:grid-cols-3">
              {/* Limited Card */}
              <button
                type="button"
                onClick={() => setLogLevel(LogLevel.LIMITED)}
                className={`group relative text-left rounded-2xl p-6 border transition-all duration-300 cursor-pointer flex flex-col justify-between glass-panel ${
                  logLevel === LogLevel.LIMITED
                    ? "border-violet-500 bg-violet-950/20 shadow-lg shadow-violet-500/10 scale-[1.01]"
                    : "border-zinc-800 bg-zinc-900/40 hover:border-zinc-700"
                }`}
              >
                <div className="space-y-3">
                  <div className="flex items-center justify-between">
                    <span className="font-semibold text-white tracking-wide text-lg">
                      Limited
                    </span>
                    <span
                      className={`h-4 w-4 rounded-full border flex items-center justify-center ${
                        logLevel === LogLevel.LIMITED
                          ? "border-violet-400 bg-violet-500"
                          : "border-zinc-600"
                      }`}
                    >
                      {logLevel === LogLevel.LIMITED && (
                        <span className="h-2 w-2 rounded-full bg-white" />
                      )}
                    </span>
                  </div>
                  <p className="text-xs text-zinc-400 leading-relaxed">
                    Usage statistics are collected to monitor service health and
                    improve performance. No document content or transcriptions
                    are stored.
                  </p>
                </div>
              </button>

              {/* Transcription Card */}
              <button
                type="button"
                onClick={() => setLogLevel(LogLevel.TRANSCRIPTION)}
                className={`group relative text-left rounded-2xl p-6 border transition-all duration-300 cursor-pointer flex flex-col justify-between glass-panel ${
                  logLevel === LogLevel.TRANSCRIPTION
                    ? "border-violet-500 bg-violet-950/20 shadow-lg shadow-violet-500/10 scale-[1.01]"
                    : "border-zinc-800 bg-zinc-900/40 hover:border-zinc-700"
                }`}
              >
                <div className="space-y-3">
                  <div className="flex items-center justify-between">
                    <span className="font-semibold text-white tracking-wide text-lg">
                      Transcription
                    </span>
                    <span
                      className={`h-4 w-4 rounded-full border flex items-center justify-center ${
                        logLevel === LogLevel.TRANSCRIPTION
                          ? "border-violet-400 bg-violet-500"
                          : "border-zinc-600"
                      }`}
                    >
                      {logLevel === LogLevel.TRANSCRIPTION && (
                        <span className="h-2 w-2 rounded-full bg-white" />
                      )}
                    </span>
                  </div>
                  <p className="text-xs text-zinc-400 leading-relaxed">
                    The service processes documents ephemerally; original files
                    are immediately discarded and never used for research. For
                    operational diagnostics, the system retains only high-level
                    processing data, such as image transcriptions and extracted
                    text fragments.
                  </p>
                </div>
              </button>

              {/* Full Card */}
              <button
                type="button"
                onClick={() => setLogLevel(LogLevel.FULL)}
                className={`group relative text-left rounded-2xl p-6 border transition-all duration-300 cursor-pointer flex flex-col justify-between glass-panel ${
                  logLevel === LogLevel.FULL
                    ? "border-violet-500 bg-violet-950/20 shadow-lg shadow-violet-500/10 scale-[1.01]"
                    : "border-zinc-800 bg-zinc-900/40 hover:border-zinc-700"
                }`}
              >
                <div className="space-y-3">
                  <div className="flex items-center justify-between">
                    <span className="font-semibold text-white tracking-wide text-lg">
                      Full
                    </span>
                    <span
                      className={`h-4 w-4 rounded-full border flex items-center justify-center ${
                        logLevel === LogLevel.FULL
                          ? "border-violet-400 bg-violet-500"
                          : "border-zinc-600"
                      }`}
                    >
                      {logLevel === LogLevel.FULL && (
                        <span className="h-2 w-2 rounded-full bg-white" />
                      )}
                    </span>
                  </div>
                  <p className="text-xs text-zinc-400 leading-relaxed">
                    Allows the service to utilize transcription data and source
                    material exclusively for academic research and
                    service-enhancement analytics.
                  </p>
                </div>
              </button>
            </div>

            <div className="flex justify-center pt-4">
              <button
                id="tos-confirm-btn"
                onClick={() => setStatus("idle")}
                className="relative group overflow-hidden px-8 py-4 rounded-xl bg-gradient-to-r from-violet-600 to-indigo-600 text-white font-semibold shadow-lg shadow-violet-600/20 hover:shadow-violet-600/40 hover:scale-[1.02] transition-all duration-200"
              >
                <span className="absolute inset-0 bg-white/10 opacity-0 group-hover:opacity-100 transition-opacity duration-300" />
                <span className="flex items-center gap-2">
                  Confirm & Continue
                  <svg
                    className="h-4 w-4 transform group-hover:translate-x-1 transition-transform"
                    fill="none"
                    stroke="currentColor"
                    viewBox="0 0 24 24"
                    xmlns="http://www.w3.org/2000/svg"
                  >
                    <path
                      strokeLinecap="round"
                      strokeLinejoin="round"
                      strokeWidth={2}
                      d="M14 5l7 7m0 0l-7 7m7-7H3"
                    />
                  </svg>
                </span>
              </button>
            </div>
          </div>
        )}

        {/* State 1: Idle (Upload prompted) */}
        {status === "idle" && (
          <div className="space-y-8 animate-fade-in">
            <div className="text-center space-y-3">
              <h1 className="text-4xl font-extrabold tracking-tight text-white sm:text-5xl">
                Automatic Alt Text Embedder{" "}
                {/*
                <span className="bg-gradient-to-r from-violet-400 via-fuchsia-400 to-indigo-400 bg-clip-text text-transparent">
                  Accessibility
                </span>*/}
              </h1>
              <p className="text-zinc-400 max-w-lg mx-auto text-base">
                Upload your Word document (.docx) or PowerPoint presentation (.pptx) to automatically generate
                relevant alt text for images in the document.
              </p>
            </div>

            {/* Drag & Drop Zone */}
            <div
              id="dropzone"
              onClick={triggerFileSelect}
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onDrop={handleDrop}
              className={`group relative rounded-2xl p-10 border-2 border-dashed transition-all duration-300 cursor-pointer flex flex-col items-center justify-center text-center glass-panel glass-panel-hover ${
                isDragOver
                  ? "border-violet-500 bg-violet-950/20 scale-[1.01]"
                  : "border-zinc-800 bg-zinc-900/40"
              }`}
            >
              {/* Pulsing glow background internally */}
              <div className="absolute inset-0 rounded-2xl bg-gradient-to-tr from-violet-600/5 to-indigo-600/5 opacity-0 group-hover:opacity-100 transition-opacity duration-500 pointer-events-none" />

              <input
                id="file-upload-input"
                type="file"
                ref={fileInputRef}
                onChange={handleFileChange}
                accept=".docx,.pptx"
                className="hidden"
              />

              {/* Icon */}
              <div className="relative mb-6 p-4 rounded-full bg-zinc-900 border border-zinc-800 text-zinc-400 group-hover:text-violet-400 group-hover:border-violet-500/30 group-hover:bg-violet-950/10 transition-all duration-300">
                <svg
                  className="h-8 w-8"
                  fill="none"
                  stroke="currentColor"
                  viewBox="0 0 24 24"
                  xmlns="http://www.w3.org/2000/svg"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    strokeWidth={1.5}
                    d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"
                  />
                </svg>
              </div>

              {file ? (
                <div className="space-y-2">
                  <p className="text-lg font-semibold text-white tracking-wide truncate max-w-md">
                    {file.name}
                  </p>
                  <p className="text-sm text-zinc-500 font-mono">
                    {formatBytes(file.size)}
                  </p>
                  <span className="inline-block mt-3 text-xs bg-violet-500/10 text-violet-400 px-3 py-1 rounded-full border border-violet-500/20 font-medium">
                    Ready to process
                  </span>
                </div>
              ) : (
                <div className="space-y-2">
                  <p className="text-lg font-medium text-zinc-300">
                    Drag and drop your file here, or{" "}
                    <span className="text-violet-400 hover:text-violet-300 underline underline-offset-4 font-semibold">
                      browse
                    </span>
                  </p>
                  <p className="text-xs text-zinc-500">
                    Supports Microsoft Word (.docx) and PowerPoint (.pptx) files up to 25MB
                  </p>
                </div>
              )}
            </div>

            {/* Action Button */}
            {file && (
              <div className="flex justify-center pt-2">
                <button
                  id="process-file-btn"
                  onClick={startProcessing}
                  className="relative group overflow-hidden px-8 py-4 rounded-xl bg-gradient-to-r from-violet-600 to-indigo-600 text-white font-semibold shadow-lg shadow-violet-600/20 hover:shadow-violet-600/40 hover:scale-[1.02] transition-all duration-200"
                >
                  {/* Subtle hover gradient shine */}
                  <span className="absolute inset-0 bg-white/10 opacity-0 group-hover:opacity-100 transition-opacity duration-300" />
                  <span className="flex items-center gap-2">
                    Start Processing
                    <svg
                      className="h-4 w-4 transform group-hover:translate-x-1 transition-transform"
                      fill="none"
                      stroke="currentColor"
                      viewBox="0 0 24 24"
                      xmlns="http://www.w3.org/2000/svg"
                    >
                      <path
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        strokeWidth={2}
                        d="M14 5l7 7m0 0l-7 7m7-7H3"
                      />
                    </svg>
                  </span>
                </button>
              </div>
            )}
          </div>
        )}

        {/* State 2 & 3: Loading Screen (Uploading / Processing) */}
        {(status === "uploading" || status === "processing") && (
          <div className="glass-panel rounded-2xl p-12 text-center space-y-8 relative overflow-hidden">
            {/* Spinning Radar Ring */}
            <div className="flex justify-center">
              <div className="relative h-24 w-24">
                {/* Outer animated ring */}
                <div className="absolute inset-0 rounded-full border-4 border-violet-500/10 border-t-violet-500 animate-spin" />
                {/* Inner pulsing pulse */}
                <div className="absolute inset-4 rounded-full bg-violet-600/20 border border-violet-500/30 flex items-center justify-center animate-pulse-glow">
                  <svg
                    className="h-6 w-6 text-violet-400"
                    fill="none"
                    stroke="currentColor"
                    viewBox="0 0 24 24"
                    xmlns="http://www.w3.org/2000/svg"
                  >
                    <path
                      strokeLinecap="round"
                      strokeLinejoin="round"
                      strokeWidth={2}
                      d="M4 4v5h.582m15.356 2A8.001 8.001 0 1121.21 8H18"
                    />
                  </svg>
                </div>
              </div>
            </div>
            <div></div>
          </div>
        )}

        {/* State 4: Success & Download Screen */}
        {status === "success" && processedFile && (
          <div className="space-y-8 animate-fade-in">
            <div className="text-center space-y-3">
              <div className="mx-auto h-16 w-16 rounded-full bg-emerald-500/10 border border-emerald-500/30 flex items-center justify-center text-emerald-400 mb-4 shadow-lg shadow-emerald-500/10">
                <svg
                  className="h-8 w-8"
                  fill="none"
                  stroke="currentColor"
                  viewBox="0 0 24 24"
                  xmlns="http://www.w3.org/2000/svg"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    strokeWidth={2.5}
                    d="M5 13l4 4L19 7"
                  />
                </svg>
              </div>
              <h1 className="text-3xl font-extrabold tracking-tight text-white">
                Alt Text Generation Complete
              </h1>
              <p className="text-zinc-400 max-w-md mx-auto text-sm">
                Your file has been processed successfully and alt text has been
                added to all images. The result is available for download for 5
                minutes.
              </p>
            </div>

            {/* File details panel */}
            <div className="glass-panel rounded-2xl p-6 flex items-center justify-between gap-4 border border-zinc-800 bg-zinc-900/20">
              <div className="flex items-center gap-4 min-w-0">
                <div className="p-3 bg-zinc-900 rounded-xl border border-zinc-800 text-violet-400 flex-shrink-0">
                  <svg
                    className="h-8 w-8"
                    fill="none"
                    stroke="currentColor"
                    viewBox="0 0 24 24"
                    xmlns="http://www.w3.org/2000/svg"
                  >
                    <path
                      strokeLinecap="round"
                      strokeLinejoin="round"
                      strokeWidth={1.5}
                      d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"
                    />
                  </svg>
                </div>
                <div className="min-w-0">
                  <p className="text-sm font-semibold text-white truncate max-w-sm">
                    {processedFile.name}
                  </p>
                  <p className="text-xs text-zinc-500 font-mono mt-0.5">
                    {formatBytes(processedFile.size)} • {processedFile.name.endsWith(".docx") ? "Microsoft Word Document" : "PowerPoint Presentation"}
                  </p>
                </div>
              </div>
              <span className="text-[10px] font-bold uppercase tracking-wider bg-emerald-500/10 text-emerald-400 border border-emerald-500/20 px-2.5 py-1 rounded-md flex-shrink-0">
                Complete
              </span>
            </div>

            {/* Action buttons */}
            <div className="flex flex-col sm:flex-row gap-4 justify-center items-center">
              <button
                id="download-file-btn"
                onClick={downloadProcessedFile}
                className="w-full sm:w-auto relative group overflow-hidden px-8 py-4 rounded-xl bg-gradient-to-r from-violet-600 to-indigo-600 text-white font-semibold shadow-lg shadow-violet-600/20 hover:shadow-violet-600/40 hover:scale-[1.02] transition-all duration-200"
              >
                <span className="absolute inset-0 bg-white/10 opacity-0 group-hover:opacity-100 transition-opacity duration-300" />
                <span className="flex items-center justify-center gap-2">
                  <svg
                    className="h-5 w-5"
                    fill="none"
                    stroke="currentColor"
                    viewBox="0 0 24 24"
                    xmlns="http://www.w3.org/2000/svg"
                  >
                    <path
                      strokeLinecap="round"
                      strokeLinejoin="round"
                      strokeWidth={2}
                      d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"
                    />
                  </svg>
                  Download {processedFile.name.endsWith(".docx") ? "DOCX" : "PPTX"}
                </span>
              </button>

              <button
                id="reset-btn"
                onClick={resetState}
                className="w-full sm:w-auto px-8 py-4 rounded-xl bg-zinc-900 border border-zinc-800 text-zinc-300 font-semibold hover:bg-zinc-800 hover:text-white hover:border-zinc-700 transition-all duration-200"
              >
                Process Another Document
              </button>
            </div>
          </div>
        )}

        {/* State 5: Error Screen */}
        {status === "error" && (
          <div className="space-y-6 animate-fade-in">
            <div className="glass-panel border-red-500/20 bg-red-950/5 rounded-2xl p-8 text-center space-y-4">
              <div className="mx-auto h-12 w-12 rounded-full bg-red-500/10 border border-red-500/30 flex items-center justify-center text-red-400 mb-2">
                <svg
                  className="h-6 w-6"
                  fill="none"
                  stroke="currentColor"
                  viewBox="0 0 24 24"
                  xmlns="http://www.w3.org/2000/svg"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    strokeWidth={2}
                    d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z"
                  />
                </svg>
              </div>
              <h3 className="text-lg font-semibold text-white">
                Processing Error
              </h3>
              <p className="text-zinc-400 text-sm max-w-md mx-auto leading-relaxed">
                {errorMsg ||
                  "An unexpected error occurred while processing the document. Please check the file and try again."}
              </p>
            </div>

            <div className="flex justify-center">
              <button
                id="error-reset-btn"
                onClick={resetState}
                className="px-6 py-3 rounded-xl bg-zinc-900 border border-zinc-800 text-zinc-300 font-medium hover:bg-zinc-800 hover:text-white transition-all duration-200"
              >
                Go Back & Try Again
              </button>
            </div>
          </div>
        )}
      </main>

      {/* Footer */}
      <footer className="relative w-full max-w-5xl px-6 mt-16 text-center text-xs text-zinc-600 z-10">
        <p>
          © {new Date().getFullYear()} EduAlly Accessibility Engine. All rights
          reserved.
        </p>
      </footer>
    </div>
  );
}
