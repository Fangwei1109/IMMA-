const assert = require("assert");
const test = require("node:test");
const { getLlmProviders, getResolvedLlmConfig, resolveLlmProvider } = require("./llm");

test("DeepSeek is selected before OpenAI and Alibaba when configured", () => {
  const originalDeepSeekKey = process.env.DEEPSEEK_API_KEY;
  const originalDeepSeekModel = process.env.DEEPSEEK_MODEL;
  const originalOpenAiKey = process.env.OPENAI_API_KEY;
  const originalAliKey = process.env.ALI_API_KEY;

  process.env.DEEPSEEK_API_KEY = "test-deepseek-key";
  process.env.DEEPSEEK_MODEL = "deepseek-v4-pro";
  process.env.OPENAI_API_KEY = "test-openai-key";
  process.env.ALI_API_KEY = "test-ali-key";

  assert.equal(resolveLlmProvider("auto"), "deepseek");

  const providers = getLlmProviders();
  assert.equal(providers.defaultProvider, "deepseek");
  assert.ok(providers.providers.find((provider) => provider.id === "deepseek")?.configured);

  const config = getResolvedLlmConfig("deepseek");
  assert.equal(config.providerId, "deepseek");
  assert.equal(config.model, "deepseek-v4-pro");

  restoreEnv("DEEPSEEK_API_KEY", originalDeepSeekKey);
  restoreEnv("DEEPSEEK_MODEL", originalDeepSeekModel);
  restoreEnv("OPENAI_API_KEY", originalOpenAiKey);
  restoreEnv("ALI_API_KEY", originalAliKey);
});

function restoreEnv(key, value) {
  if (value === undefined) {
    delete process.env[key];
    return;
  }

  process.env[key] = value;
}
