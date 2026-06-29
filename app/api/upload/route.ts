import { NextRequest, NextResponse } from "next/server";

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get("uploadedFile") as File;

    if (!file) {
      return NextResponse.json({ error: "No file uploaded" }, { status: 400 });
    }

    const forwardData = new FormData();
    forwardData.append("uploadedFile", file, file.name);
    forwardData.append("logLevel", formData.get("logLevel") as string);

    const apiKey = process.env.BACKEND_API_KEY;
    const isPptx = file.name.endsWith(".pptx");
    const apiAddr = isPptx ? process.env.BACKEND_PPTX_API_ADDR : process.env.BACKEND_API_ADDR;

    if (!apiAddr) {
      return NextResponse.json(
        { error: `Backend API address is not configured for ${isPptx ? ".pptx" : ".docx"} files` },
        { status: 500 }
      );
    }

    const response = await fetch(apiAddr, {
      method: "POST",
      headers: {
        ...(apiKey ? { "Authorization": `Bearer ${apiKey}` } : {}),
      },
      body: forwardData,
    });

    if (!response.ok) {
      const errorText = await response.text();
      return NextResponse.json(
        { error: `Backend API returned status ${response.status}: ${errorText}` },
        { status: response.status }
      );
    }

    const json = await response.json();
    return NextResponse.json(json);
  } catch (error: unknown) {
    console.error("Upload error:", error);
    const errorMessage = error instanceof Error ? error.message : "Failed to process file upload";
    return NextResponse.json(
      { error: errorMessage },
      { status: 500 }
    );
  }
}
