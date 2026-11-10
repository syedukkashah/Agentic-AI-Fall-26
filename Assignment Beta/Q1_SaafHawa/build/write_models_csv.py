"""Writes data/models.csv (run from the Q1 folder: python build/write_models_csv.py data/models.csv). Every value was read from the source IDs in the *_src columns (see data/sources.csv)."""
import csv
import sys

NR = "not reported"
cols = ["model", "side", "provider", "variant_used", "params_total_B", "params_active_B", "context_tokens", "max_output_tokens", "spec_src",
        "price_in", "price_out", "price_cached", "price_cache_write", "host", "quantization", "price_src",
        "cache_applies_1500", "cache_min_tokens", "cache_ttl", "cache_src",
        "rpm", "itpm", "otpm", "rate_note", "rate_src",
        "ttft_s", "ttft_variant", "output_tps", "speed_variant", "speed_src",
        "bench1_name", "bench1_score", "bench2_tool_name", "bench2_tool_score", "bench_src",
        "tool_calling", "json_output", "reasoning_settings", "reasoning_tokens_at_setting", "licence", "commercial_self_host", "feature_src"]

NO_CAP = "no OpenRouter platform cap for paid models; upstream provider limits apply but are not published"

rows = [
    # ---------------- closed ----------------
    ["GPT-6 Luna", "closed", "OpenAI", "reasoning_effort=none", "not disclosed", "not disclosed", 1050000, 128000, "S3",
     0.10, 0.50, 0.01, 0.125, "OpenAI", "n/a", "S1", True, 1024, "30 min (default)", "S2",
     5000, 2000000, "", "Build tier (lowest listed); TPM is combined", "S3",
     1.15, "non-reasoning", 101, "max (AA model page)", "S21;S22",
     "AA Intelligence Index v5.0.1 (non-reasoning)", 9, "tool-use benchmark", NR, "S21",
     "yes", "yes", "none, low, medium, high, xhigh, max", "0 (none)", "proprietary", "no", "S3;S19"],
    ["GPT-6.1 Terra", "closed", "OpenAI", "reasoning_effort=none", "not disclosed", "not disclosed", 1050000, 128000, "S4",
     1.00, 5.00, 0.05, 1.25, "OpenAI", "n/a", "S1", True, 1024, "30 min (default)", "S2",
     5000, 1000000, "", "Build tier (lowest listed); TPM is combined", "S4",
     1.63, "non-reasoning", 74, "max (AA model page)", "S21;S22",
     "AA Intelligence Index v5.0.1 (non-reasoning)", 31, "tool-use benchmark", NR, "S21",
     "yes", "yes", "none, low, medium, high, xhigh, max", "0 (none)", "proprietary", "no", "S4;S19"],
    ["Claude Haiku 5.5", "closed", "Anthropic", "effort=low (no non-thinking mode)", "not disclosed", "not disclosed", 1000000, 128000, "S6",
     0.10, 0.50, 0.01, 0.125, "Anthropic", "n/a", "S5", True, 512, "5 min (default; 1 h optional)", "S7",
     4000, 2000000, 400000, "Start tier; cache reads excluded from ITPM", "S8",
     6.79, "low", 232, "max (AA model page)", "S21;S22",
     "AA Intelligence Index v5.0.1 (low)", 19, "tool-use benchmark", NR, "S21",
     "yes", "yes", "adaptive thinking; effort low, medium (default), high, xhigh, max", NR, "proprietary", "no", "S6;S9;S19"],
    ["Claude Sonnet 5.5", "closed", "Anthropic", "effort=low (no non-thinking mode)", "not disclosed", "not disclosed", 1000000, 128000, "S6",
     2.00, 10.00, 0.10, 2.50, "Anthropic", "n/a", "S5", True, 512, "5 min (default; 1 h optional)", "S7",
     4000, 2000000, 400000, "Start tier; cache reads excluded from ITPM", "S8",
     2.40, "low", 151, "max (AA model page)", "S21;S22",
     "AA Intelligence Index v5.0.1 (low)", 26, "tool-use benchmark", NR, "S21",
     "yes", "yes", "adaptive thinking; effort low, medium, high (default), xhigh, max", NR, "proprietary", "no", "S6;S9;S19"],
    ["Gemini 3.5 Flash-Lite", "closed", "Google", "thinking_level=minimal (default)", "not disclosed", "not disclosed", 1048576, 65536, "S18",
     0.30, 2.50, 0.03, "storage $1.00/M tok/h", "Google", "n/a", "S10", False, "not listed (other Gemini 3.x: 4,096)", "implicit; TTL not stated", "S11",
     NR, NR, "", "Tier 1 per-model limits only visible inside a signed-in AI Studio account", "S12",
     4.08, "default (minimal)", 292, "default (AA model page)", "S21;S22",
     "AA Intelligence Index v5.0.1", 12, "tool-use benchmark", NR, "S21",
     "yes", "yes", "minimal (default), low, medium, high", NR, "proprietary", "no", "S13;S19"],
    ["Gemini 3.8 Flash", "closed", "Google", "thinking_level=low", "not disclosed", "not disclosed", 1048576, 65536, "S18",
     0.75, 3.75, 0.075, "storage $0.50/M tok/h", "Google", "n/a", "S10", False, 4096, "implicit; TTL not stated", "S11",
     NR, NR, "", "Tier 1 per-model limits only visible inside a signed-in AI Studio account; prices are promotional until 2026-12-31 ($1.50/$7.50 from 2027-01-01)", "S12;S10",
     3.74, "low", 111, "high (AA model page)", "S21;S22",
     "AA Intelligence Index v5.0.1 (low)", 23, "tool-use benchmark", NR, "S21",
     "yes", "yes", "low, medium (default), high", NR, "proprietary", "no", "S13;S19"],
    ["Grok 4.20 0309 (non-reasoning)", "closed", "xAI", "non-reasoning model id", "not disclosed", "not disclosed", 1000000, "not stated", "S14",
     1.25, 2.50, 0.20, "not documented", "xAI", "n/a", "S14", False, "not documented (caching page 404)", "not documented", "S15",
     1776, 10000000, "", "Tier 0 (new account): 37 RPS and 10M TPM; RPM derived as 37 x 48 from the RPS = RPM/48 rule on the same page", "S15",
     NR, "AA page: model deprecated, not re-benchmarked", NR, "", "S22",
     "IFBench (reasoning variant)", 82.9, "tool-use benchmark", NR, "S23",
     "yes", "yes", "separate reasoning / non-reasoning model ids", "0 (non-reasoning)", "proprietary", "no", "S14;S19"],
    # ---------------- open-weight ----------------
    ["Ministral 3 14B", "open", "Mistral AI", "instruct (non-reasoning)", 13.9, "dense", 262144, 209715, "S24;S19",
     0.20, 0.20, 0.02, "none", "Mistral", "not stated", "S16;S19", True, 64, "not stated", "S17",
     NR, NR, "", "no OpenRouter platform cap for paid models; Mistral tier page returned 404", "S20",
     2.54, "non-reasoning", 65, "non-reasoning (AA model page)", "S21;S22",
     "GPQA Diamond (vendor; reasoning-variant table)", 71.2, "tool-use benchmark", NR, "S24",
     "yes", "yes", "none in Instruct (separate Reasoning checkpoint)", "0", "Apache 2.0", "yes", "S24;S19"],
    ["Mistral Small 4", "open", "Mistral AI", "reasoning off", 119.4, 6.5, 262144, 209715, "S25;S19",
     0.15, 0.60, 0.015, "none", "Mistral", "not stated", "S16;S19", True, 64, "not stated", "S17",
     NR, NR, "", "no OpenRouter platform cap for paid models; Mistral tier page returned 404", "S20",
     1.77, "non-reasoning", 147, "reasoning (AA model page)", "S21;S22",
     "AA Intelligence Index v5.0.1 (reasoning)", 4, "tool-use benchmark", NR, "S22",
     "yes", "yes", "reasoning effort configurable per request", "0 (off)", "Apache 2.0", "yes", "S25;S19"],
    ["Qwen3.8 27B", "open", "Alibaba (Qwen)", "thinking off", 27.8, "dense (hybrid linear/full attention)", 262144, 65536, "S26;S19",
     0.051, 0.362, 0.006, "none", "PrimeIntellect (via OpenRouter)", "fp8", "S19", False, "not published by host", "not published", "S19",
     NR, NR, "", NO_CAP, "S20",
     5.31, "non-reasoning", 45, "xhigh (AA model page)", "S21;S22",
     "AA Intelligence Index v5.0.1 (non-reasoning)", 10, "Terminal Bench 2.1 (agentic; vendor)", 73.0, "S21;S26",
     "yes", "yes", "thinking on / off", "0 (off)", "Apache 2.0", "yes", "S26;S19"],
    ["Gemma 4 31B", "open", "Google", "thinking off", 30.7, "dense (sliding + full attention)", 262144, 235929, "S27;S19",
     0.10, 0.34, 0.10, "none", "CoreWeave (via OpenRouter)", "fp4", "S19", False, "not published by host", "not published", "S19",
     NR, NR, "", NO_CAP, "S20",
     3.90, "non-reasoning", 36, "reasoning (AA model page)", "S21;S22",
     "GPQA Diamond (vendor)", 84.3, "Tau2, average of 3 (vendor)", 76.9, "S27",
     "yes", "yes", "thinking on / off", "0 (off)", "Apache 2.0", "yes", "S27;S19"],
    ["gpt-oss-120b", "open", "OpenAI", "reasoning=low (lowest)", 116.8, 5.1, 131072, 117964, "S28;S19",
     0.030, 0.150, 0.015, "none", "Near AI (via OpenRouter)", "fp4", "S19", False, "not published by host", "not published", "S19",
     NR, NR, "", NO_CAP, "S20",
     2.20, "low", 144, "high (AA model page)", "S21;S22",
     "benchmark", NR, "tool-use benchmark", NR, "S22",
     "yes", "yes", "low, medium, high", NR, "Apache 2.0", "yes", "S28;S19"],
    ["DeepSeek V4.1 Flash 1017", "open", "DeepSeek", "non-reasoning", 763.2, "not stated (MoE, 6 of 384 experts per token)", 1048576, 131072, "S29;S31;S19",
     0.025, 0.250, 0.0015, "none", "Particle AI (via OpenRouter)", "not stated", "S19", False, "not published by host", "not published", "S19",
     NR, NR, "", NO_CAP, "S20",
     1.32, "non-reasoning", 144, "max (AA model page)", "S21;S22",
     "AA Intelligence Index v5.0.1 (non-reasoning)", 19, "tool-use benchmark", NR, "S21",
     "yes", "yes", "non-reasoning, high, max", "0 (non-reasoning)", "MIT", "yes", "S29;S19"],
    ["GLM-5.3-Flash", "open", "Z.ai", "default (reasoning; no non-reasoning variant benchmarked)", 320, 18, 1048576, 131072, "S30;S19",
     0.034, 0.083, 0.030, "none", "Relace (via OpenRouter)", "fp4", "S19", False, "not published by host", "not published", "S19",
     NR, NR, "", NO_CAP, "S20",
     94.96, "default (reasoning)", 49, "default (AA model page)", "S21;S22",
     "AA Intelligence Index v5.0.1", 28, "tool-use benchmark", NR, "S21",
     "yes", "yes", "reasoning (default)", NR, "MIT", "yes", "S30;S19"],
]

assert all(len(r) == len(cols) for r in rows), [len(r) for r in rows]

# Tool-use benchmarks read from the Artificial Analysis leaderboard with the Intelligence Index expanded into its
# component evaluations (source S33; values captured by screenshot on 2026-11-09).
# model -> (variant row used, AutomationBench-AA %, Terminal-Bench 4.0 %)
AA_TOOL = {
    "GPT-6 Luna": ("non-reasoning", 9.1, 1.5),
    "GPT-6.1 Terra": ("non-reasoning", 51.0, 26.8),
    "Claude Haiku 5.5": ("low", 22.9, 12.6),
    "Claude Sonnet 5.5": ("low", 49.4, 20.7),
    "Gemini 3.5 Flash-Lite": ("default", 25.0, 1.0),
    "Gemini 3.8 Flash": ("low", 36.5, 10.1),
    "Ministral 3 14B": ("default", 0.6, 0.0),
    "Mistral Small 4": ("reasoning; non-reasoning row not evaluated", 1.0, 0.0),
    "Qwen3.8 27B": ("non-reasoning", 12.4, 0.0),
    "Gemma 4 31B": ("reasoning (AA default)", 4.9, 0.0),
    "gpt-oss-120b": ("high; low row not evaluated", 0.2, 0.0),
    "DeepSeek V4.1 Flash 1017": ("non-reasoning", 48.8, 9.1),
    "GLM-5.3-Flash": ("default", 60.4, 32.8),
}
# vendor-reported tool-use scores (model cards) kept as notes; not comparable with AA's independent runs
VENDOR_TOOL = {
    "Qwen3.8 27B": "Terminal Bench 2.1 = 73.0 (vendor, S26)",
    "Gemma 4 31B": "Tau2 avg of 3 = 76.9% (vendor, S27)",
}
cols += ["terminal_bench4_score", "vendor_tool_note"]
i_name, i_score, i_src = cols.index("bench2_tool_name"), cols.index("bench2_tool_score"), cols.index("bench_src")
for r in rows:
    if r[0] in AA_TOOL:
        variant, automation, terminal = AA_TOOL[r[0]]
        r[i_name], r[i_score] = f"AutomationBench-AA ({variant})", automation
        r[i_src] = r[i_src] + ";S33"
        r.append(terminal)
    else:
        r[i_name] = "AutomationBench-AA (not benchmarked by AA: deprecated model)"
        r.append(NR)
    r.append(VENDOR_TOOL.get(r[0], ""))
with open(sys.argv[1], "w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    writer.writerow(cols)
    writer.writerows(rows)
print("wrote", sys.argv[1], ":", len(rows), "models,", len(cols), "columns")
