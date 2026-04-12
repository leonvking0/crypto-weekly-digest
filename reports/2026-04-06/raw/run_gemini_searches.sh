#!/bin/bash
set -euo pipefail
RAW_DIR="$1"
run_query() {
  local idx="$1"
  local slug="$2"
  local prompt="$3"
  local outfile="$RAW_DIR/${idx}-${slug}.md"
  echo "[$(date -u +%FT%TZ)] Running $idx-$slug" >&2
  if gemini -p "$prompt" > "$outfile" 2>"$outfile.err"; then
    :
  else
    if grep -qi 'rate limit\|quota\|429' "$outfile.err"; then
      echo "Rate limited on $slug, retrying with 10s wait" >&2
      sleep 10
      if ! gemini -p "$prompt" > "$outfile" 2>>"$outfile.err"; then
        echo "Falling back to gemini-3-pro-preview for $slug" >&2
        gemini -m gemini-3-pro-preview -p "$prompt" > "$outfile" 2>>"$outfile.err"
      fi
    else
      echo "Primary attempt failed for $slug, falling back to gemini-3-pro-preview" >&2
      gemini -m gemini-3-pro-preview -p "$prompt" > "$outfile" 2>>"$outfile.err"
    fi
  fi
  sleep 3
}

run_query 01 market-overview "Search the web for the most important Market Overview & Prices crypto news from the past 7 days. Focus on Messari and The Block as primary sources but include other major outlets. List the top 3-5 developments with brief descriptions, dates, and source URLs where available. Be factual and specific, no speculation. Cover BTC/ETH/SOL and major altcoin price action, market cap changes, Fear & Greed index context, key technical levels, and weekly performance summary."
run_query 02 layer1-layer2 "Search the web for the most important Layer 1 & Layer 2 crypto news from the past 7 days. Focus on Messari and The Block as primary sources but include other major outlets. List the top 3-5 developments with brief descriptions, dates, and source URLs where available. Be factual and specific, no speculation. Cover major chain upgrades, launches, and TVL shifts."
run_query 03 defi "Search the web for the most important DeFi crypto news from the past 7 days. Focus on Messari and The Block as primary sources but include other major outlets. List the top 3-5 developments with brief descriptions, dates, and source URLs where available. Be factual and specific, no speculation. Cover protocol launches, governance, yield developments, and TVL movements."
run_query 04 perp-dexes "Search the web for the most important Perpetual DEX crypto news from the past 7 days. Focus on Messari and The Block as primary sources but include other major outlets. List the top 3-5 developments with brief descriptions, dates, and source URLs where available. Be factual and specific, no speculation. Cover volume records, new platforms, and market share shifts."
run_query 05 stablecoins "Search the web for the most important Stablecoins crypto news from the past 7 days. Focus on Messari and The Block as primary sources but include other major outlets. List the top 3-5 developments with brief descriptions, dates, and source URLs where available. Be factual and specific, no speculation. Cover supply changes, new issuers, and regulatory actions."
run_query 06 ai-crypto "Search the web for the most important AI x Crypto news from the past 7 days. Focus on Messari and The Block as primary sources but include other major outlets. List the top 3-5 developments with brief descriptions, dates, and source URLs where available. Be factual and specific, no speculation. Cover AI agent tokens, compute networks, and AI-integrated protocols."
run_query 07 prediction-markets "Search the web for the most important Prediction Markets crypto news from the past 7 days. Focus on Messari and The Block as primary sources but include other major outlets. List the top 3-5 developments with brief descriptions, dates, and source URLs where available. Be factual and specific, no speculation. Cover volume, new markets, and platform developments."
run_query 08 regulation-policy "Search the web for the most important Regulation & Policy crypto news from the past 7 days. Focus on Messari and The Block as primary sources but include other major outlets. List the top 3-5 developments with brief descriptions, dates, and source URLs where available. Be factual and specific, no speculation. Cover SEC/CFTC actions, legislation, and country-level moves."
run_query 09 fundraising-ma "Search the web for the most important Fundraising & M&A crypto news from the past 7 days. Focus on Messari and The Block as primary sources but include other major outlets. List the top 3-5 developments with brief descriptions, dates, and source URLs where available. Be factual and specific, no speculation. Cover notable raises, acquisitions, and token launches."
run_query 10 security-incidents "Search the web for the most important Security Incidents crypto news from the past 7 days. Focus on Messari and The Block as primary sources but include other major outlets. List the top 3-5 developments with brief descriptions, dates, and source URLs where available. Be factual and specific, no speculation. Cover hacks, exploits, rug pulls, and fund recoveries with dollar amounts."
run_query 11 upcoming-catalysts "Search the web for the most important Upcoming Catalysts in crypto for the next 1-2 weeks. Focus on Messari and The Block as primary sources but include other major outlets. List the top 5 upcoming events with dates, affected assets, and why they matter. Be factual and specific, no speculation. Cover token unlocks, mainnet launches, governance votes, conferences, ETF decisions, and protocol upgrades."
run_query 12 ecosystem-infra "Search the web for the most important Ecosystem & Infrastructure crypto news from the past 7 days. Focus on Messari and The Block as primary sources but include other major outlets. List the top 3-5 developments with brief descriptions, dates, and source URLs where available. Be factual and specific, no speculation. Cover wallets, bridges, oracles, and tooling."
run_query 13 synthesis "What were the single biggest crypto industry stories of the past 7 days? Give me the top 5 most impactful events across all sectors. Prioritize Messari and The Block coverage. Include dates and source URLs where available."
