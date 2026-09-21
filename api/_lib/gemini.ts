import { GoogleGenAI, Type } from "@google/genai";

let aiClient: GoogleGenAI | null = null;
export function getGenAI(): GoogleGenAI {
  if (!aiClient) {
    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) {
      throw new Error(
        "GEMINI_API_KEY environment variable is missing. Please configure GEMINI_API_KEY in your Vercel or hosting environment settings."
      );
    }
    aiClient = new GoogleGenAI({
      apiKey,
      httpOptions: {
        headers: {
          "User-Agent": "aistudio-build",
        },
      },
    });
  }
  return aiClient;
}

// Role-based System Instructions for Maritime Ship Chandler Operations in Bangladesh
export const ROLE_SYSTEM_INSTRUCTIONS: Record<string, string> = {
  maritime_analyst: `You are the Senior Maritime Operations & Ship Chandler Specialist for Zainee Enterprise, a government-licensed ship chandler and marine contractor operating across all major ports of Bangladesh:
- Chittagong (Chattogram) Port (CPA jetties, Karnaphuli river berths, and Outer Anchorage Alpha/Bravo/Charlie, Kutubdia lighterage).
- Mongla Port (Pashur river & Harbaria anchorage).
- Matarbari Deep Sea Port & Payra Port.

You have comprehensive knowledge of the BANGLADESHI MARITIME & LOCAL SUPPLY MARKET:
1. SOURCING HUBS & CURRENT BANGLADESH PRICING:
   - Fresh & Dry Provisions (Khatunganj & Chaktai wholesale commodity market, Chittagong):
     * Fresh Halal Beef: ~750-850 BDT/kg
     * Fresh Mutton: ~1,100-1,250 BDT/kg
     * Fresh Broiler Chicken: ~180-220 BDT/kg
     * Farm Fresh Eggs: ~130-150 BDT/dozen
     * Premium Basmati Rice: ~140-180 BDT/kg; Miniket: ~75-85 BDT/kg
     * Seasonal Vegetables (Potatoes, Onions, Cabbage, Tomatoes): ~45-65 BDT/kg
     * Bottled Mineral Water (1.5L x 12 cases): ~220-250 BDT/case
     * Bulk Fresh Drinking Water via Supply Barge to Outer Anchorage: ~2,500-3,500 BDT/metric ton
   - Engine & Deck Marine Stores (Sadarghat, Strand Road & Agrabad, Chittagong):
     * 220m 24mm/28mm 8-Strand Polypropylene Mooring Rope: ~58,000-82,000 BDT/coil
     * Galvanized Wire Ropes (6x36 WS IWRC): ~360-550 BDT/meter
     * SOLAS Approved Adult Lifejackets with light & whistle: ~2,900-3,900 BDT/pc
     * SOLAS Immersion Suits (MED certified): ~12,000-16,000 BDT/pc
     * Zinc Anodes for hull & ballast tanks: ~950-1,150 BDT/kg
     * Industrial Cotton Cleaning Rags/Waste: ~65-80 BDT/kg
   - Ship Breaking Yard OEM Machinery & Spares (Bhatiary & Sitakunda, Chittagong):
     * Genuine reconditioned marine valves (JIS / DIN bronze globe, gate, storm valves), pump impellers, purifiers (Alfa Laval/Mitsubishi), diesel engine spares (Daihatsu, Yanmar, MAN B&W) available at 40-70% savings compared to new imports.

2. SPEED & ACCURACY DIRECTIVES:
   - Provide answers FAST, accurately, and straight to the point.
   - Always quote current estimated prices in Bangladeshi Taka (BDT / Taka).
   - Format item lists clearly in markdown tables with: Item Description, IMPA/ISSA Code (if applicable), Standard Unit, and Estimated Unit Price in Taka (BDT).
   - Specify whether the pricing is for Port Berth delivery or Outer Anchorage delivery (including supply boat / launch lighterage costs: ~35,000-55,000 BDT per launch trip from Ghat 15).`,

  pricing_negotiator: `You are the Commercial Director & Senior Pricing Strategist for Zainee Enterprise in Bangladesh.
Your expertise covers:
- Current Bangladeshi ship supply commercial margins (typically 15%-25% on provisions, 20%-35% on technical deck/engine stores, 30%-50% on reconditioned Sitakunda shipyard spares).
- Outer Anchorage lighterage launch boat hire tariffs (~35,000-55,000 BDT/trip from Sadarghat/Ghat 15 to Chittagong Outer Anchorage depending on sea state and waiting time).
- Local maritime supply pricing and invoicing in Bangladeshi Taka (BDT / Taka).
- Payment terms in shipping: Cash on Delivery (COD), Cash Against Documents (CAD), 30-day DA with foreign owners, and Master's General Receipt.
- National Board of Revenue (NBR) Bangladesh VAT guidelines: supplies to foreign-flagged vessels under bonded customs supply enjoy export zero-rated VAT status.
Give fast, sharp, numbers-driven advice that maximizes profit margin while offering competitive quotes to ship owners and vessel managers.`,

  vessel_auditor: `You are the Quality Control & Maritime Compliance Auditor for Zainee Enterprise, Chittagong, Bangladesh.
Your duties are:
- Auditing quotations, challans, and commercial invoices against Chittagong Port Authority (CPA) and international shipping standards.
- Verifying complete vessel details (Vessel Name, Call Sign, IMO Number, Port/Berth, Requisition No, and Purchase Order).
- Ensuring quantities, standard marine units (PCS, KGS, COIL, MTR, SET, LTR, DRUM), unit prices, and extended amounts calculate with zero mathematical discrepancies.
- Validating delivery documentation for Master / Chief Engineer / Chief Officer sign-off and ship stamp.
Highlight missing vessel info, unit mismatches, or pricing errors in a clean, high-speed checklist format.`,
};

export interface ExtractItemsParams {
  text: string;
  model?: string;
}

export async function executeExtractItems({ text, model }: ExtractItemsParams) {
  const chosenModel = model === "gemini-3.1-flash-lite" ? "gemini-3.1-flash-lite" : "gemini-3.8-flash";
  const candidateModels = [chosenModel, "gemini-3.1-flash-lite", "gemini-3.8-flash"].filter(
    (m, idx, arr) => arr.indexOf(m) === idx
  );

  const ai = getGenAI();
  let response: any = null;

  for (const m of candidateModels) {
    try {
      response = await ai.models.generateContent({
        model: m,
        contents: `Extract all maritime supply items, equipment, provisions, and spare parts from this request or text into structured items:
\n\n"""\n${text}\n"""`,
        config: {
          systemInstruction: `You are an expert maritime item parser for ship chandlers operating at Chittagong & Mongla ports in Bangladesh. 
Analyze the input text (which might be an email, RFQ, WhatsApp list, or captain's requisition).
Extract each distinct line item with its description (including sizes, dimensions, or IMPA/ISSA code if present), quantity (as a string or number), standard unit (e.g. PCS, KGS, SET, MTR, LTR, BDL, PKT, BOX, PAIR, ROLL, COIL, DRUM), and price (if mentioned or suggested for Bangladeshi market, otherwise empty string "").
Return an array of JSON objects matching the schema with high speed and zero extra narrative.`,
          responseMimeType: "application/json",
          responseSchema: {
            type: Type.ARRAY,
            items: {
              type: Type.OBJECT,
              properties: {
                desc: { type: Type.STRING, description: "Detailed item name, spec, or IMPA code" },
                qty: { type: Type.STRING, description: "Quantity requested" },
                unit: { type: Type.STRING, description: "Standard maritime unit of measurement" },
                price: { type: Type.STRING, description: "Unit price if given or suggested, else empty string" },
              },
              required: ["desc", "qty", "unit"],
            },
          },
        },
      });
      if (response) break;
    } catch (err: any) {
      console.warn(`Extraction failed on model ${m} (${err?.message}), trying fallback...`);
    }
  }

  if (!response) {
    throw new Error("Unable to extract items at this moment.");
  }

  let items = [];
  try {
    items = JSON.parse(response.text || "[]");
  } catch (parseError) {
    console.error("JSON parse error on extracted items:", parseError);
    items = [];
  }

  return { items };
}

export interface AuditDocumentParams {
  document: any;
  model?: string;
}

export async function executeAuditDocument({ document, model }: AuditDocumentParams) {
  const chosenModel = model === "gemini-3.1-pro-preview" ? "gemini-3.1-pro-preview" : "gemini-3.8-flash";
  const ai = getGenAI();

  const prompt = `Conduct a rigorous, professional maritime commercial audit on this document for Zainee Enterprise, Chittagong Port, Bangladesh:
Document Type: ${document.docType}
Challan/Invoice/Quotation No: ${document.challanNo || document.invoiceNo || "N/A"}
Date: ${document.date}
Client / Messers: ${document.messers || "MISSING"}
Vessel Name: ${document.vesselName || "MISSING"}
Port / Address: ${document.address || "MISSING"}
Currency: ${document.currency}
Grand Total: ${document.grandTotal}
Transportation/Boat Hire: ${document.transportationFee || 0}
Rows count: ${document.rows?.length || 0}
Items sample:
${JSON.stringify(document.rows?.slice(0, 30) || [], null, 2)}

Provide:
1. Overall Audit Score (0-100)
2. Critical Issues / Missing Mandatory Shipping Fields (e.g. Missing Vessel name, zero prices, missing units)
3. Pricing & Calculation Verification
4. Recommendations for immediate correction
5. Recommended professional cover remark for the Ship Master or Superintendent.`;

  const response = await ai.models.generateContent({
    model: chosenModel,
    contents: prompt,
    config: {
      systemInstruction: ROLE_SYSTEM_INSTRUCTIONS.vessel_auditor,
      temperature: 0.4,
    },
  });

  return {
    auditResult: response.text || "",
    modelUsed: chosenModel,
  };
}
