import matplotlib
matplotlib.use("Agg")

import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

df = pd.read_csv("crypto_clean.csv")

print(f"Total coins loaded: {len(df)}")
print(f"Columns: {list(df.columns)}")

df_filtered = df[(df["Price (USD)"] >= 0) & (df["Price (USD)"] <= 5)].copy()
print(f"\nCoins with price $0-$5: {len(df_filtered)}")


df_filtered["Price 1h Ago"] = df_filtered["Price (USD)"] / (1 + df_filtered["1h Change (%)"] / 100)
df_filtered["Price 7d Ago"]  = df_filtered["Price (USD)"] / (1 + df_filtered["7d Change (%)"] / 100)
df_filtered["Price 24h Ago"] = df_filtered["Price (USD)"] / (1 - df_filtered["24h Change (%)"] / 100)

top10 = df_filtered.nlargest(10, "Price 1h Ago").reset_index(drop=True)

print("\n" + "=" * 62)
print("  TOP 10 CRYPTO COINS (by 1-Hour-Before Price, $0-$5 range)")
print("=" * 62)
print(top10[["Coin Name", "Symbol", "Price (USD)", "Price 1h Ago",
             "Price 24h Ago", "Price 7d Ago"]].to_string(index=False))
print()

coin_labels = top10["Coin Name"] + "\n(" + top10["Symbol"] + ")"
x = np.arange(len(coin_labels))
bar_width = 0.35

fig, ax = plt.subplots(figsize=(16, 9))

bars_7d  = ax.bar(x - bar_width/2, top10["Price 7d Ago"],  bar_width,
                  label="7 Days Before Price", color="#6366f1", edgecolor="white", linewidth=0.6)
bars_24h = ax.bar(x + bar_width/2, top10["Price 24h Ago"], bar_width,
                  label="24 Hours Before Price", color="#f59e0b", edgecolor="white", linewidth=0.6)

for bar in bars_7d:
    height = bar.get_height()
    ax.annotate(f"${height:.4f}",
                xy=(bar.get_x() + bar.get_width() / 2, height),
                xytext=(0, 5), textcoords="offset points",
                ha="center", va="bottom", fontsize=8, fontweight="bold", color="#6366f1")

for bar in bars_24h:
    height = bar.get_height()
    ax.annotate(f"${height:.4f}",
                xy=(bar.get_x() + bar.get_width() / 2, height),
                xytext=(0, 5), textcoords="offset points",
                ha="center", va="bottom", fontsize=8, fontweight="bold", color="#b45309")

ax.scatter(x, top10["Price (USD)"], color="#ef4444", zorder=5, s=60, marker="D", label="Current Price")
for i, price in enumerate(top10["Price (USD)"]):
    ax.annotate(f"${price:.4f}", xy=(x[i], price),
                xytext=(0, 10), textcoords="offset points",
                ha="center", fontsize=7, color="#ef4444", fontstyle="italic")

ax.set_xlabel("Coin Name", fontsize=12, fontweight="bold", labelpad=10)
ax.set_ylabel("Price (USD)", fontsize=12, fontweight="bold", labelpad=10)
ax.set_title("Top 10 Crypto Coins ($0-$5 Range) - 7-Day & 24-Hour Before Prices\n"
             "(Ranked by 1-Hour Before Price)",
             fontsize=14, fontweight="bold", pad=15)
ax.set_xticks(x)
ax.set_xticklabels(coin_labels, fontsize=9, ha="center")
ax.legend(fontsize=10, loc="upper right")
ax.grid(axis="y", alpha=0.3, linestyle="--")
ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)

plt.tight_layout()
plt.savefig("top10_crypto_chart.png", dpi=150, bbox_inches="tight")
print("\nChart saved as 'top10_crypto_chart.png'")
