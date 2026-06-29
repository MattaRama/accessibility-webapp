import { NextRequest, NextResponse } from "next/server";
import http from "http";

// Helper function to send a GET request with a body using Node.js native http module
function getWithBody(
  url: string,
  body: string,
  headers: Record<string, string>
): Promise<{ status: number; text: string }> {
  return new Promise((resolve, reject) => {
    const parsedUrl = new URL(url);
    const options = {
      hostname: parsedUrl.hostname,
      port: parsedUrl.port || 80,
      path: parsedUrl.pathname + parsedUrl.search,
      method: "GET",
      headers: {
        ...headers,
        "Content-Length": Buffer.byteLength(body),
      },
    };

    const req = http.request(options, (res) => {
      let data = "";
      res.on("data", (chunk) => {
        data += chunk;
      });
      res.on("end", () => {
        resolve({
          status: res.statusCode || 200,
          text: data,
        });
      });
    });

    req.on("error", (e) => {
      reject(e);
    });

    req.write(body);
    req.end();
  });
}

export async function GET(request: NextRequest) {
  try {
    const searchParams = request.nextUrl.searchParams;
    const jobId = searchParams.get("jobId");

    if (!jobId) {
      return NextResponse.json({ error: "jobId is required" }, { status: 400 });
    }

    const apiKey = process.env.BACKEND_API_KEY;

    const requestBody = JSON.stringify({ jobId });
    const headers: Record<string, string> = {
      "Content-Type": "application/json",
    };
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
    const result = await getWithBody(
      `${baseAddr}subscribe`,
      requestBody,
      headers
    );

    if (result.status < 200 || result.status >= 300) {
      return NextResponse.json(
        { error: `Backend subscription failed (status ${result.status}): ${result.text}` },
        { status: result.status }
      );
    }

    const json = JSON.parse(result.text);
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
