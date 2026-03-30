#!/bin/bash
set -e
OUT=".tmp-research"
mkdir -p "$OUT"

run_search() {
  local name="$1"
  local query="$2"
  echo "=== Searching: $name ==="
  gemini -p "$query" > "$OUT/$name.txt" 2>&1 || true
  sleep 4
}

run_search "01-market" "Search the web for the most important crypto market overview news from March 23-29, 2026. BTC/ETH/SOL price action, key technical levels, weekly performance. Focus on Messari, The Block, CoinDesk. List top 3-5 developments with dates and sources."

run_search "02-l1l2" "Search the web for the most important Layer 1 and Layer 2 crypto news from March 23-29, 2026. Major chain upgrades, launches, TVL shifts. Focus on Messari, The Block. List top 3-5 developments with dates and sources."

run_search "03-defi" "Search the web for the most important DeFi news from March 23-29, 2026. Protocol launches, governance, yield developments, TVL movements. Focus on Messari, The Block. List top 3-5 developments with dates and sources."

run_search "04-perpdex" "Search the web for the most important perpetual DEX news from March 23-29, 2026. Volume records, new platforms, market share shifts, Hyperliquid dYdX GMX. Focus on Messari, The Block. List top 3-5 developments."

run_search "05-stablecoins" "Search the web for the most important stablecoin news from March 23-29, 2026. Supply changes, new issuers, regulatory actions, USDT USDC. Focus on Messari, The Block. List top 3-5 developments."

run_search "06-aicrypto" "Search the web for the most important AI x Crypto news from March 23-29, 2026. AI agent tokens, compute networks, AI-integrated protocols. Focus on Messari, The Block. List top 3-5 developments."

run_search "07-prediction" "Search the web for the most important prediction market crypto news from March 23-29, 2026. Polymarket, volume, new markets. Focus on Messari, The Block. List top 3-5 developments."

run_search "08-regulation" "Search the web for the most important crypto regulation and policy news from March 23-29, 2026. SEC CFTC actions, legislation, country-level moves. Focus on Messari, The Block. List top 3-5 developments."

run_search "09-fundraising" "Search the web for the most important crypto fundraising and M&A news from March 23-29, 2026. Notable raises, acquisitions, token launches. Focus on Messari, The Block. List top 3-5 developments."

run_search "10-security" "Search the web for the most important crypto security incidents from March 23-29, 2026. Hacks, exploits, rug pulls, fund recoveries with dollar amounts. Focus on Messari, The Block. List all incidents."

run_search "11-catalysts" "Search the web for upcoming crypto catalysts in the next 1-2 weeks after March 29, 2026. Token unlocks, mainnet launches, governance votes, conferences, ETF decisions, protocol upgrades. Be specific with dates."

run_search "12-ecosystem" "Search the web for the most important crypto ecosystem and infrastructure news from March 23-29, 2026. Wallets, bridges, oracles, tooling developments. Focus on Messari, The Block. List top 3-5 developments."

run_search "13-synthesis" "What were the single biggest crypto industry stories of March 23-29, 2026? Give me the top 5 most impactful events across all sectors. Prioritize Messari and The Block coverage. Be factual and specific."

echo "=== ALL SEARCHES COMPLETE ==="
