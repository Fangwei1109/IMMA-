const DEFAULT_TAVILY_URL = "https://api.tavily.com/search";
const DEFAULT_SERPAPI_URL = "https://serpapi.com/search.json";

function getResearchProviders() {
  return {
    defaultProvider: resolveResearchProvider(),
    providers: [
      {
        id: "none",
        configured: true,
        label: "Disabled",
      },
      {
        id: "tavily",
        configured: Boolean(process.env.TAVILY_API_KEY),
        label: "Tavily",
      },
      {
        id: "serpapi",
        configured: Boolean(process.env.SERPAPI_API_KEY),
        label: "SerpAPI",
      },
    ],
  };
}

function resolveResearchProvider(requestedProvider = "auto") {
  if (requestedProvider && requestedProvider !== "auto" && requestedProvider !== "none") {
    return requestedProvider;
  }

  if (process.env.TAVILY_API_KEY) {
    return "tavily";
  }

  if (process.env.SERPAPI_API_KEY) {
    return "serpapi";
  }

  return "none";
}

async function enrichCompanyContext({
  company,
  meetingType,
  templateId,
  transcriptText,
  materialText,
  focusAreas = [],
  providerId = "auto",
}) {
  const resolvedProvider = resolveResearchProvider(providerId);

  if (!company || resolvedProvider === "none") {
    return null;
  }

  const queries = buildQueries({
    company,
    meetingType,
    templateId,
    transcriptText,
    materialText,
    focusAreas,
  });

  if (queries.length === 0) {
    return null;
  }

  if (resolvedProvider === "tavily") {
    return searchWithTavily(queries);
  }

  if (resolvedProvider === "serpapi") {
    return searchWithSerpApi(queries);
  }

  return null;
}

function buildQueries({ company, meetingType, templateId, transcriptText, materialText, focusAreas = [] }) {
  const context = `${transcriptText || ""}\n${materialText || ""}`.toLowerCase();
  const queries = [
    `${company} official website company profile`,
  ];

  const focusSet = new Set(focusAreas);

  if (focusSet.size === 0 || focusSet.has("team")) {
    queries.push(`${company} founder management team background`);
  }

  if (focusSet.size === 0 || focusSet.has("funding") || focusSet.has("ipo")) {
    queries.push(`${company} funding valuation founders`);
  }

  if (focusSet.size === 0 || focusSet.has("commercialization") || focusSet.has("traction")) {
    queries.push(`${company} customers partnerships deployments`);
  }

  if (focusSet.has("product")) {
    queries.push(`${company} product technology roadmap`);
  }

  if (templateId === "weekly-report" || /pre-ipo|fundraising|ipo|listing/.test(context)) {
    queries.push(`${company} IPO fundraising valuation`);
  }

  if (/autonomous|robot|delivery|logistics|vehicle|无人|自动驾驶/.test(context)) {
    queries.push(`${company} autonomous logistics deployment customers`);
  }

  if (meetingType && meetingType !== "general") {
    queries.push(`${company} ${meetingType} latest news`);
  }

  return [...new Set(queries)].slice(0, 5);
}

async function searchWithTavily(queries) {
  const apiKey = process.env.TAVILY_API_KEY;
  if (!apiKey) {
    return null;
  }

  const results = [];
  for (const query of queries) {
    const response = await fetch(process.env.TAVILY_API_URL || DEFAULT_TAVILY_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        api_key: apiKey,
        query,
        search_depth: "advanced",
        max_results: 5,
        include_answer: true,
      }),
    });

    if (!response.ok) {
      const detail = await response.text();
      throw new Error(`Tavily search failed: ${detail}`);
    }

    const payload = await response.json();
    results.push({
      query,
      answer: payload.answer || "",
      results: (payload.results || []).map((item) => ({
        title: item.title,
        url: item.url,
        content: item.content,
      })),
    });
  }

  return {
    providerId: "tavily",
    searches: results,
  };
}

async function searchWithSerpApi(queries) {
  const apiKey = process.env.SERPAPI_API_KEY;
  if (!apiKey) {
    return null;
  }

  const results = [];
  for (const query of queries) {
    const searchParams = new URLSearchParams({
      engine: "google",
      q: query,
      api_key: apiKey,
      num: "5",
    });

    const response = await fetch(`${process.env.SERPAPI_URL || DEFAULT_SERPAPI_URL}?${searchParams.toString()}`);
    if (!response.ok) {
      const detail = await response.text();
      throw new Error(`SerpAPI search failed: ${detail}`);
    }

    const payload = await response.json();
    results.push({
      query,
      answer: payload.answer_box?.answer || payload.answer_box?.snippet || "",
      results: (payload.organic_results || []).slice(0, 5).map((item) => ({
        title: item.title,
        url: item.link,
        content: item.snippet,
      })),
    });
  }

  return {
    providerId: "serpapi",
    searches: results,
  };
}

module.exports = {
  enrichCompanyContext,
  getResearchProviders,
  resolveResearchProvider,
};
