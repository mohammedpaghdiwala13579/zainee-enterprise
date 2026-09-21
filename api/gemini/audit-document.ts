import { executeAuditDocument } from "../_lib/gemini.js";

export default async function handler(req: any, res: any) {
  // Enable CORS
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");

  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed. Use POST." });
  }

  try {
    const body = typeof req.body === "string" ? JSON.parse(req.body) : req.body || {};
    const { document, model } = body;

    if (!document || !Array.isArray(document.rows)) {
      return res.status(400).json({ error: "Invalid document payload for audit." });
    }

    const result = await executeAuditDocument({ document, model });
    return res.status(200).json(result);
  } catch (error: any) {
    console.error("Error in Vercel api/gemini/audit-document:", error);
    return res.status(500).json({
      error: error?.message || "Failed to audit document with Gemini AI.",
    });
  }
}
