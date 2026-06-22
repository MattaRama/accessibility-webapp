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
    const apiAddr = process.env.BACKEND_API_ADDR;
    const response = await fetch(apiAddr!, {
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
  } catch (error: any) {
    console.error("Upload error:", error);
    return NextResponse.json(
      { error: error.message || "Failed to process file upload" },
      { status: 500 }
    );
  }
}
