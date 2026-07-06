import { NextRequest, NextResponse } from "next/server";

export async function GET(request: NextRequest) {
  try {
    const searchParams = request.nextUrl.searchParams;
    const jobId = searchParams.get("jobId");

    if (!jobId) {
      return NextResponse.json({ error: "jobId is required" }, { status: 400 });
    }

    const apiKey = process.env.BACKEND_API_KEY;

    const headers: Record<string, string> = {};
    if (apiKey) {
      headers["Authorization"] = `Bearer ${apiKey}`;
    }

    const fileType = searchParams.get("fileType");
    const isPptx = fileType === "pptx";
    const apiAddr = isPptx ? process.env.BACKEND_PPTX_API_ADDR : process.env.BACKEND_API_ADDR;

    if (!apiAddr) {
      return NextResponse.json(
        { error: `Backend API address is not configured for ${isPptx ? ".pptx" : ".docx"} files` },
        { status: 500 }
      );
    }

    const baseAddr = apiAddr.endsWith("/") ? apiAddr : `${apiAddr}/`;
    const response = await fetch(`${baseAddr}subscribe?jobId=${jobId}`, {
      method: "GET",
      headers,
    });

    if (!response.ok) {
      const errorText = await response.text();
      return NextResponse.json(
        { error: `Backend subscription failed (status ${response.status}): ${errorText}` },
        { status: response.status }
      );
    }

    const json = await response.json();
    return NextResponse.json(json);
  } catch (error: unknown) {
    console.error("Subscription error:", error);
    const errorMessage = error instanceof Error ? error.message : "Failed to subscribe to job";
    return NextResponse.json(
      { error: errorMessage },
      { status: 500 }
    );
  }
}
