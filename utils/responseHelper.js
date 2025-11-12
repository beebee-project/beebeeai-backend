function cleanAIResponse(response) {
  if (response.startsWith("```") && response.endsWith("```")) {
    const lines = response.split("\n");

    let cleaned = lines.slice(1, -1).join("\n").trim();

    if (
      cleaned.toLowerCase().startsWith("excel") ||
      cleaned.toLowerCase().startsWith("json")
    ) {
      cleaned = cleaned.substring(cleaned.indexOf("\n") + 1).trim();
    }
    return cleaned;
  }

  return response.trim();
}

module.exports = { cleanAIResponse };
