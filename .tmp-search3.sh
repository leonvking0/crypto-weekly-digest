#!/bin/bash
OUT=".tmp-research"
run_search() {
  echo "=== Searching: $1 ==="
  gemini -p "$2" > "$OUT/$1.txt" 2>&1 || true
  sleep 4
}

run_search "07-prediction" "Search the web for the most important prediction market crypto news from March 23-29, 2026. Polymarket, volume, new markets. List top 3-5 developments with dates and sources."

run_search "08-regulation" "Search the web for the most important crypto regulation and policy news from March 23-29, 2026. SEC CFTC actions, legislation. List top 3-5 developments."

run_search "09-fundraising" "Search the web for the most important crypto fundraising and M&A news from March 23-29, 2026. Notable raises, acquisitions. List top 3-5."

run_search "10-security" "Search the web for the most important crypto security incidents from March 23-29, 2026. Hacks, exploits with dollar amounts. List all."

run_search "11-catalysts" "Search the web for upcoming crypto catalysts in next 1-2 weeks after March 29, 2026. Token unlocks, mainnet launches, conferences. Be specific with dates."

run_search "12-ecosystem" "Search the web for the most important crypto ecosystem and infrastructure news from March 23-29, 2026. Wallets, bridges, oracles. List top 3-5."

run_search "13-synthesis" "What were the single biggest crypto industry stories of March 23-29, 2026? Top 5 most impactful events across all sectors. Be factual and specific."

echo "=== ALL DONE ==="
