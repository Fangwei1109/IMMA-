const DEFAULT_OPENAI_BASE_URL = "https://api.openai.com/v1";
const DEFAULT_DEEPSEEK_BASE_URL = "https://api.deepseek.com";
const DEFAULT_ALI_BASE_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1";

function readEnv(name) {
  return String(process.env[name] || "").trim();
}

function getLlmProviders() {
  return {
    defaultProvider: resolveLlmProvider(),
    providers: [
      {
        id: "auto",
        label: "Auto",
        configured: true,
      },
      {
        id: "openai",
        label: "OpenAI",
        configured: Boolean(readEnv("OPENAI_API_KEY")),
        model: readEnv("OPENAI_MODEL") || "gpt-4.1-mini",
      },
      {
        id: "deepseek",
        label: "DeepSeek",
        configured: Boolean(readEnv("DEEPSEEK_API_KEY")),
        model: readEnv("DEEPSEEK_MODEL") || "deepseek-v4-flash",
      },
      {
        id: "ali",
        label: "Ali / Qwen",
        configured: Boolean(readEnv("ALI_API_KEY")),
        model: readEnv("ALI_MODEL") || "qwen-plus",
      },
      {
        id: "none",
        label: "Disabled",
        configured: true,
      },
    ],
  };
}

function resolveLlmProvider(requestedProvider = "auto") {
  if (requestedProvider && requestedProvider !== "auto") {
    return requestedProvider;
  }

  if (readEnv("DEEPSEEK_API_KEY")) {
    return "deepseek";
  }

  if (readEnv("OPENAI_API_KEY")) {
    return "openai";
  }

  if (readEnv("ALI_API_KEY")) {
    return "ali";
  }

  return "none";
}

function getResolvedLlmConfig(requestedProvider = "auto") {
  const providerId = resolveLlmProvider(requestedProvider);

  if (providerId === "openai") {
    return {
      providerId,
      apiKey: readEnv("OPENAI_API_KEY"),
      baseUrl: readEnv("OPENAI_BASE_URL") || DEFAULT_OPENAI_BASE_URL,
      model: readEnv("OPENAI_MODEL") || "gpt-4.1-mini",
    };
  }

  if (providerId === "deepseek") {
    return {
      providerId,
      apiKey: readEnv("DEEPSEEK_API_KEY"),
      baseUrl: readEnv("DEEPSEEK_BASE_URL") || DEFAULT_DEEPSEEK_BASE_URL,
      model: readEnv("DEEPSEEK_MODEL") || "deepseek-v4-flash",
    };
  }

  if (providerId === "ali") {
    return {
      providerId,
      apiKey: readEnv("ALI_API_KEY"),
      baseUrl: readEnv("ALI_BASE_URL") || DEFAULT_ALI_BASE_URL,
      model: readEnv("ALI_MODEL") || "qwen-plus",
    };
  }

  return {
    providerId: "none",
    apiKey: "",
    baseUrl: "",
    model: "",
  };
}

async function extractImageMaterialText({ requestedProvider = "auto", mimeType, buffer, fileName = "image" }) {
  const config = getResolvedLlmConfig(requestedProvider);

  if (!config.apiKey || config.providerId === "none") {
    return {
      providerId: config.providerId,
      text: "",
      usedVision: false,
    };
  }

  const base64 = Buffer.from(buffer).toString("base64");
  const response = await fetch(`${config.baseUrl}/chat/completions`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${config.apiKey}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      model: config.model,
      temperature: 0.1,
      messages: [
        {
          role: "system",
          content:
            "You extract factual text and structured investor-relevant facts from uploaded images. Do OCR where possible. Return concise plain text only. Do not invent facts that are not visible.",
        },
        {
          role: "user",
          content: [
            {
              type: "text",
              text: [
                `File name: ${fileName}`,
                "Please extract any visible text, labels, tables, metrics, dates, team names, fundraising information, roadmap items, product specs, or commercialization signals from this image.",
                "Format:",
                "Visible text:",
                "- ...",
                "",
                "Structured notes:",
                "- ...",
              ].join("\n"),
            },
            {
              type: "image_url",
              image_url: {
                url: `data:${mimeType || "image/png"};base64,${base64}`,
              },
            },
          ],
        },
      ],
    }),
  });

  if (!response.ok) {
    const detail = await response.text();
    throw new Error(`Image extraction failed: ${detail}`);
  }

  const payload = await response.json();
  const content = payload.choices?.[0]?.message?.content;

  return {
    providerId: config.providerId,
    text: typeof content === "string" ? content.trim() : "",
    usedVision: true,
  };
}

module.exports = {
  extractImageMaterialText,
  getLlmProviders,
  resolveLlmProvider,
  getResolvedLlmConfig,
};
