from __future__ import annotations

import json
import math
import os
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd


REPO_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_DATA_PATH = Path(r"C:\Users\Larry.Nie\Downloads\B50_X_COMMENT.xlsx")
FIGURES_DIR = REPO_ROOT / "outputs" / "figures"
TABLES_DIR = REPO_ROOT / "outputs" / "tables"
DOCS_DIR = REPO_ROOT / "docs"

ATTENTION_COL = "Author's attention quantity"
REACTION_COLS = ["retweets count", "likes", "Comment views"]

TRUMP_TERMS = ["trump", "president", "potus", "maga", "realdonaldtrump"]
POSITIVE_STANCE_TERMS = [
    "greatest",
    "elite",
    "legendary",
    "epic",
    "awesome",
    "incredible",
    "good golfer",
    "incredible golfer",
    "insane",
    "insanely",
    "national treasure",
    "goat",
    "goats",
    "lfg",
    "stripes the driver",
    "has game",
    "can play",
    "flusher",
    "best round",
    "fun to watch",
]
NEGATIVE_STANCE_TERMS = [
    "felon",
    "racist",
    "rapist",
    "pedo",
    "simp",
    "disgusted",
    "cheat",
    "cheats",
    "adjudicated",
    "convicted",
    "asshole",
    "clown",
    "gtfooh",
    "suckers and losers",
    "dangerous",
    "worst person",
    "disappointed",
    "staged his own assassination",
    "sex offender",
    "sexual assault",
    "be better",
    "boycott",
    "polarizing person",
    "alienating",
]
DEPOLITICIZING_TERMS = [
    "no politics",
    "politics out of it",
    "love of golf",
    "healing",
    "play together",
    "very little politicin",
    "say what you want about politics",
    "pleasant surprise",
    "fun and lighthearted",
    "lighthearted dialogue",
]
POLITICAL_FRAME_TERMS = [
    "trump",
    "president",
    "potus",
    "maga",
    "politic",
    "political",
    "election",
    "felon",
    "racist",
    "democrat",
    "republican",
    "america",
    "45",
]
SPORT_FRAME_TERMS = [
    "golf",
    "bryson",
    "break 50",
    "driver",
    "swing",
    "course",
    "round",
    "putt",
    "youtube",
    "club",
    "golfer",
    "break 47",
    "shot",
    "hole",
    "fairway",
]
STANCE_ORDER = [
    "Trump-referential/unclear",
    "pro-Trump/supportive",
    "anti-Trump/oppositional",
    "depoliticizing/bridge",
    "sport-centered",
    "other/unclear",
]
FRAME_ORDER = [
    "political",
    "blended sport-politics",
    "sport",
    "depoliticizing bridge",
    "other",
]


def ensure_dirs() -> None:
    FIGURES_DIR.mkdir(parents=True, exist_ok=True)
    TABLES_DIR.mkdir(parents=True, exist_ok=True)
    DOCS_DIR.mkdir(parents=True, exist_ok=True)


def resolve_data_path() -> Path:
    raw = os.environ.get("BREAK50_DATA_PATH")
    data_path = Path(raw) if raw else DEFAULT_DATA_PATH
    if not data_path.exists():
        raise FileNotFoundError(
            "Could not find the local dataset. Set BREAK50_DATA_PATH to the path of B50_X_COMMENT.xlsx."
        )
    return data_path


def count_hits(text: str, terms: list[str]) -> int:
    return sum(term in text for term in terms)


def has_any(text: str, terms: list[str]) -> bool:
    return any(term in text for term in terms)


def classify_comment(text: str) -> tuple[str, str, int, int]:
    lower = text.lower()
    pos_hits = count_hits(lower, POSITIVE_STANCE_TERMS)
    neg_hits = count_hits(lower, NEGATIVE_STANCE_TERMS)
    depol_hits = count_hits(lower, DEPOLITICIZING_TERMS)
    has_trump = has_any(lower, TRUMP_TERMS)
    has_political = has_any(lower, POLITICAL_FRAME_TERMS)
    has_sport = has_any(lower, SPORT_FRAME_TERMS)

    if neg_hits > 0 and neg_hits >= pos_hits:
        stance = "anti-Trump/oppositional"
    elif depol_hits > 0:
        stance = "depoliticizing/bridge"
    elif has_trump and pos_hits > 0:
        stance = "pro-Trump/supportive"
    elif has_trump:
        stance = "Trump-referential/unclear"
    elif has_sport:
        stance = "sport-centered"
    else:
        stance = "other/unclear"

    if depol_hits > 0:
        frame = "depoliticizing bridge"
    elif has_political and has_sport:
        frame = "blended sport-politics"
    elif has_political:
        frame = "political"
    elif has_sport:
        frame = "sport"
    else:
        frame = "other"

    return stance, frame, int(neg_hits > 0), int(depol_hits > 0)


def load_and_prepare(data_path: Path) -> pd.DataFrame:
    df = pd.read_excel(data_path).copy()

    numeric_cols = [
        ATTENTION_COL,
        "retweets count",
        "likes",
        "Comment word length",
        "Comment views",
        "followers",
        "number of media outlets",
        "number of author posts",
    ]
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["verified"] = df["blue_verified"].astype(int)
    df["english"] = (df["Comment language"] == "en").astype(int)
    df["any_retweet"] = (df["retweets count"] > 0).astype(int)
    df["any_like"] = (df["likes"] > 0).astype(int)
    df["log_attention"] = np.log1p(df[ATTENTION_COL])
    df["log_likes"] = np.log1p(df["likes"])
    df["log_views"] = np.log1p(df["Comment views"])
    df["log_media_outlets"] = np.log1p(df["number of media outlets"])
    df["log_author_posts"] = np.log1p(df["number of author posts"])

    coded = df["contents"].fillna("").astype(str).apply(classify_comment)
    df[["stance", "frame", "moral_condemnation", "depoliticizing_appeal"]] = pd.DataFrame(
        coded.tolist(),
        index=df.index,
    )

    df["stance"] = pd.Categorical(df["stance"], categories=STANCE_ORDER, ordered=True)
    df["frame"] = pd.Categorical(df["frame"], categories=FRAME_ORDER, ordered=True)
    return df


def fmt_value(value: object) -> str:
    if isinstance(value, (int, np.integer)):
        return f"{int(value):,}"
    if isinstance(value, (float, np.floating)):
        if math.isnan(float(value)):
            return ""
        if abs(float(value)) >= 1000:
            return f"{float(value):,.2f}"
        return f"{float(value):.3f}"
    if pd.isna(value):
        return ""
    return str(value)


def df_to_markdown(df: pd.DataFrame, index: bool = True) -> str:
    table = df.copy()
    if index:
        table = table.reset_index()
    headers = [str(col) for col in table.columns]
    lines = [
        "| " + " | ".join(headers) + " |",
        "| " + " | ".join(["---"] * len(headers)) + " |",
    ]
    for _, row in table.iterrows():
        lines.append("| " + " | ".join(fmt_value(v) for v in row.tolist()) + " |")
    return "\n".join(lines)


def write_markdown_table(df: pd.DataFrame, path: Path, index: bool = True) -> None:
    path.write_text(df_to_markdown(df, index=index) + "\n", encoding="utf-8")


def normal_p_value(z: float) -> float:
    return math.erfc(abs(z) / math.sqrt(2))


def ols_hc1(y: np.ndarray, x: np.ndarray, names: list[str]) -> tuple[pd.DataFrame, dict[str, float]]:
    x = np.asarray(x, dtype=float)
    y = np.asarray(y, dtype=float)
    keep = np.isfinite(y) & np.all(np.isfinite(x), axis=1)
    x = x[keep]
    y = y[keep]

    n, k = x.shape
    xtx = x.T @ x
    xtx_inv = np.linalg.inv(xtx)
    beta = xtx_inv @ (x.T @ y)
    residuals = y - x @ beta

    meat = np.zeros((k, k))
    for i in range(n):
        xi = x[i : i + 1].T
        meat += (residuals[i] ** 2) * (xi @ xi.T)
    robust_cov = (n / (n - k)) * (xtx_inv @ meat @ xtx_inv)
    robust_se = np.sqrt(np.diag(robust_cov))
    z_values = beta / robust_se
    p_values = np.array([normal_p_value(z) for z in z_values])
    ci_low = beta - 1.96 * robust_se
    ci_high = beta + 1.96 * robust_se
    rss = float(np.sum(residuals**2))
    tss = float(np.sum((y - y.mean()) ** 2))
    r2 = 1 - rss / tss if tss else float("nan")

    results = pd.DataFrame(
        {
            "term": names,
            "coef": beta,
            "robust_se": robust_se,
            "z": z_values,
            "p_value": p_values,
            "ci_low": ci_low,
            "ci_high": ci_high,
        }
    )
    return results, {"n": float(n), "r2": float(r2)}


def sample_overview(df: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "metric": [
                "Comments",
                "Unique usernames",
                "Unique comment IDs",
                "Unique source posts",
                "Date start",
                "Date end",
                "English-language comments",
                "Verified accounts",
                "Comments with zero retweets",
                "Comments mentioning Trump-related terms",
                "Comments with moral-condemnation language",
                "Comments with depoliticizing appeals",
            ],
            "value": [
                len(df),
                df["username"].nunique(),
                df["contentsid"].nunique(),
                df["Blog ID"].nunique(),
                df["date"].min().strftime("%Y-%m-%d %H:%M:%S"),
                df["date"].max().strftime("%Y-%m-%d %H:%M:%S"),
                int(df["english"].sum()),
                int(df["verified"].sum()),
                int((df["retweets count"] == 0).sum()),
                int(df["stance"].isin(["Trump-referential/unclear", "pro-Trump/supportive", "anti-Trump/oppositional", "depoliticizing/bridge"]).sum()),
                int(df["moral_condemnation"].sum()),
                int(df["depoliticizing_appeal"].sum()),
            ],
        }
    )


def descriptive_stats(df: pd.DataFrame) -> pd.DataFrame:
    cols = [
        ATTENTION_COL,
        "retweets count",
        "likes",
        "Comment word length",
        "Comment views",
        "followers",
        "number of media outlets",
        "number of author posts",
    ]
    return df[cols].agg(["count", "mean", "std", "min", "median", "max"]).T.rename(columns={"std": "sd"})


def stance_distribution(df: pd.DataFrame) -> pd.DataFrame:
    counts = df["stance"].value_counts(dropna=False).reindex(STANCE_ORDER, fill_value=0)
    return pd.DataFrame(
        {
            "stance": counts.index,
            "comments": counts.values,
            "share": counts.values / len(df),
        }
    )


def frame_distribution(df: pd.DataFrame) -> pd.DataFrame:
    counts = df["frame"].value_counts(dropna=False).reindex(FRAME_ORDER, fill_value=0)
    return pd.DataFrame(
        {
            "frame": counts.index,
            "comments": counts.values,
            "share": counts.values / len(df),
        }
    )


def engagement_by_stance(df: pd.DataFrame) -> pd.DataFrame:
    grouped = df.groupby("stance", observed=False).agg(
        comments=("stance", "size"),
        likes_mean=("likes", "mean"),
        likes_median=("likes", "median"),
        views_mean=("Comment views", "mean"),
        views_median=("Comment views", "median"),
        retweet_rate=("any_retweet", "mean"),
        attention_median=(ATTENTION_COL, "median"),
    )
    return grouped.reset_index()


def engagement_by_frame(df: pd.DataFrame) -> pd.DataFrame:
    grouped = df.groupby("frame", observed=False).agg(
        comments=("frame", "size"),
        likes_median=("likes", "median"),
        views_median=("Comment views", "median"),
        retweet_rate=("any_retweet", "mean"),
    )
    return grouped.reset_index()


def top_decile_shares(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    overall = (df["stance"].value_counts(normalize=True).reindex(STANCE_ORDER, fill_value=0)).rename("overall")
    likes_cutoff = df["likes"].quantile(0.9)
    views_cutoff = df["Comment views"].quantile(0.9)
    top_likes = df.loc[df["likes"] >= likes_cutoff, "stance"].value_counts(normalize=True).reindex(STANCE_ORDER, fill_value=0)
    top_views = df.loc[df["Comment views"] >= views_cutoff, "stance"].value_counts(normalize=True).reindex(STANCE_ORDER, fill_value=0)
    for stance in STANCE_ORDER:
        rows.append(
            {
                "stance": stance,
                "overall_share": overall[stance],
                "top_likes_share": top_likes[stance],
                "top_views_share": top_views[stance],
            }
        )
    return pd.DataFrame(rows)


def discourse_model_summary(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    categories = [
        "pro-Trump/supportive",
        "anti-Trump/oppositional",
        "depoliticizing/bridge",
        "sport-centered",
        "other/unclear",
    ]
    for category in categories:
        df[f"stance::{category}"] = (df["stance"] == category).astype(int)

    x_names = [
        "Intercept",
        "log_attention",
        "verified",
        "english",
        "log_media_outlets",
        "log_author_posts",
    ] + categories
    x = np.column_stack(
        [
            np.ones(len(df)),
            df["log_attention"],
            df["verified"],
            df["english"],
            df["log_media_outlets"],
            df["log_author_posts"],
        ]
        + [df[f"stance::{category}"] for category in categories]
    )

    outcomes = {"Likes (log1p)": "log_likes", "Comment views (log1p)": "log_views"}
    detail_frames = []
    summary_rows = []
    for label, outcome_col in outcomes.items():
        detail_df, meta = ols_hc1(df[outcome_col].to_numpy(), x, x_names)
        detail_df.insert(0, "outcome", label)
        detail_frames.append(detail_df)
        for category in categories:
            row = detail_df.loc[detail_df["term"] == category].iloc[0]
            summary_rows.append(
                {
                    "outcome": label,
                    "comparison_to_reference": f"{category} vs Trump-referential/unclear",
                    "coef": row["coef"],
                    "robust_se": row["robust_se"],
                    "ci_low": row["ci_low"],
                    "ci_high": row["ci_high"],
                    "p_value": row["p_value"],
                    "r2": meta["r2"],
                    "n": int(meta["n"]),
                }
            )
        attention_row = detail_df.loc[detail_df["term"] == "log_attention"].iloc[0]
        summary_rows.append(
            {
                "outcome": label,
                "comparison_to_reference": "log_attention control effect",
                "coef": attention_row["coef"],
                "robust_se": attention_row["robust_se"],
                "ci_low": attention_row["ci_low"],
                "ci_high": attention_row["ci_high"],
                "p_value": attention_row["p_value"],
                "r2": meta["r2"],
                "n": int(meta["n"]),
            }
        )
    return pd.concat(detail_frames, ignore_index=True), pd.DataFrame(summary_rows)


def create_stance_distribution_figure(table: pd.DataFrame) -> None:
    fig, ax = plt.subplots(figsize=(9.2, 5.4))
    bars = ax.barh(table["stance"], table["comments"], color="#2f6f8f", edgecolor="#1f2937")
    ax.set_title("Break 50 Comment Stance Distribution")
    ax.set_xlabel("Number of comments")
    ax.grid(axis="x", linestyle="--", alpha=0.3)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    for bar, value, share in zip(bars, table["comments"], table["share"]):
        ax.text(value + 8, bar.get_y() + bar.get_height() / 2, f"{value} ({share:.1%})", va="center", fontsize=9)
    fig.tight_layout()
    fig.savefig(FIGURES_DIR / "stance_distribution.png", dpi=220)
    plt.close(fig)


def create_frame_distribution_figure(table: pd.DataFrame) -> None:
    fig, ax = plt.subplots(figsize=(8.8, 4.8))
    bars = ax.bar(table["frame"], table["comments"], color="#9c6644", edgecolor="#1f2937")
    ax.set_title("Break 50 Frame Distribution")
    ax.set_ylabel("Number of comments")
    ax.grid(axis="y", linestyle="--", alpha=0.3)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.tick_params(axis="x", rotation=20)
    for bar, value, share in zip(bars, table["comments"], table["share"]):
        ax.text(bar.get_x() + bar.get_width() / 2, value + 8, f"{value}\n{share:.1%}", ha="center", va="bottom", fontsize=9)
    fig.tight_layout()
    fig.savefig(FIGURES_DIR / "frame_distribution.png", dpi=220)
    plt.close(fig)


def create_engagement_by_stance_figure(table: pd.DataFrame) -> None:
    fig, axes = plt.subplots(1, 2, figsize=(12, 4.8))
    left = table.copy()

    axes[0].barh(left["stance"], left["likes_median"], color="#508d4e", edgecolor="#1f2937")
    axes[0].set_title("Median Likes by Stance")
    axes[0].set_xlabel("Median likes")
    axes[0].grid(axis="x", linestyle="--", alpha=0.3)
    axes[0].set_axisbelow(True)

    axes[1].barh(left["stance"], left["views_median"], color="#7b2cbf", edgecolor="#1f2937")
    axes[1].set_title("Median Comment Views by Stance")
    axes[1].set_xlabel("Median views")
    axes[1].grid(axis="x", linestyle="--", alpha=0.3)
    axes[1].set_axisbelow(True)

    for ax in axes:
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)

    fig.tight_layout()
    fig.savefig(FIGURES_DIR / "engagement_by_stance.png", dpi=220)
    plt.close(fig)


def create_top_decile_figure(table: pd.DataFrame) -> None:
    fig, ax = plt.subplots(figsize=(10, 5.2))
    x = np.arange(3)
    labels = ["Overall", "Top likes decile", "Top views decile"]
    bottoms = np.zeros(3)
    colors = {
        "Trump-referential/unclear": "#4c78a8",
        "pro-Trump/supportive": "#59a14f",
        "anti-Trump/oppositional": "#e15759",
        "depoliticizing/bridge": "#b07aa1",
        "sport-centered": "#f28e2b",
        "other/unclear": "#9d9da1",
    }

    for _, row in table.iterrows():
        values = np.array([row["overall_share"], row["top_likes_share"], row["top_views_share"]])
        ax.bar(x, values, bottom=bottoms, color=colors[row["stance"]], label=row["stance"])
        bottoms += values

    ax.set_xticks(x, labels)
    ax.set_ylim(0, 1)
    ax.set_ylabel("Share of comments")
    ax.set_title("Stance Composition of Overall vs High-Engagement Comments")
    ax.legend(loc="center left", bbox_to_anchor=(1.02, 0.5), frameon=False)
    ax.grid(axis="y", linestyle="--", alpha=0.3)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    fig.tight_layout()
    fig.savefig(FIGURES_DIR / "top_decile_stance_shares.png", dpi=220)
    plt.close(fig)


def build_analysis_report(
    data_path: Path,
    df: pd.DataFrame,
    sample_table: pd.DataFrame,
    stance_table: pd.DataFrame,
    frame_table: pd.DataFrame,
    stance_engagement: pd.DataFrame,
    top_decile_table: pd.DataFrame,
    model_summary: pd.DataFrame,
) -> None:
    pro_count = int(stance_table.loc[stance_table["stance"] == "pro-Trump/supportive", "comments"].iloc[0])
    anti_count = int(stance_table.loc[stance_table["stance"] == "anti-Trump/oppositional", "comments"].iloc[0])
    bridge_count = int(stance_table.loc[stance_table["stance"] == "depoliticizing/bridge", "comments"].iloc[0])
    political_share = float(frame_table.loc[frame_table["frame"] == "political", "share"].iloc[0])
    blended_share = float(frame_table.loc[frame_table["frame"] == "blended sport-politics", "share"].iloc[0])
    sport_views = float(
        stance_engagement.loc[stance_engagement["stance"] == "sport-centered", "views_median"].iloc[0]
    )
    pro_likes = float(
        stance_engagement.loc[stance_engagement["stance"] == "pro-Trump/supportive", "likes_median"].iloc[0]
    )
    anti_views = float(
        stance_engagement.loc[stance_engagement["stance"] == "anti-Trump/oppositional", "views_median"].iloc[0]
    )
    top_likes_pro = float(top_decile_table.loc[top_decile_table["stance"] == "pro-Trump/supportive", "top_likes_share"].iloc[0])
    top_likes_anti = float(top_decile_table.loc[top_decile_table["stance"] == "anti-Trump/oppositional", "top_likes_share"].iloc[0])
    like_attention_row = model_summary[
        (model_summary["outcome"] == "Likes (log1p)")
        & (model_summary["comparison_to_reference"] == "log_attention control effect")
    ].iloc[0]
    views_attention_row = model_summary[
        (model_summary["outcome"] == "Comment views (log1p)")
        & (model_summary["comparison_to_reference"] == "log_attention control effect")
    ].iloc[0]

    report = f"""# Break 50 Analysis Report

## Overview

This revised analysis responds to the instructor feedback by shifting the project from a generic engagement-correlation study to a **P3-style contested-discourse analysis**. Instead of treating all outcomes as interchangeable engagement measures, the report asks how commenters frame Trump's appearance in Break 50, how stance positions are distributed, and which kinds of comments receive more visible reaction.

The raw Excel file remains local and outside version control. Data source used locally: `{data_path}`

## Revised Research Question

The revised study asks:

1. How is Trump's appearance in Break 50 framed in the comments: as politics, as sport, or as a blending of the two?
2. How are stance positions distributed across the discussion, specifically pro-Trump/supportive, anti-Trump/oppositional, depoliticizing/bridge, Trump-referential/unclear, sport-centered, and other/unclear comments?
3. Do visible reaction metrics such as likes and comment views cluster differently across those stance positions?

In this revised design, **author's attention quantity** is retained as a secondary control variable rather than the main research question.

## Coding Strategy

Because the dataset does not include pre-coded discourse variables, the pipeline applies a transparent keyword-based coding scheme to approximate stance and frame. The coding rules are documented in `docs/codebook.md`.

- `pro-Trump/supportive`: comments using supportive or celebratory language toward Trump or his appearance.
- `anti-Trump/oppositional`: comments using condemnation, moral critique, or rejection language.
- `depoliticizing/bridge`: comments explicitly calling for politics to be set aside in favor of golf or shared enjoyment.
- `Trump-referential/unclear`: comments that mention Trump but do not clearly resolve into support or opposition under the heuristic rules.
- `sport-centered`: comments focused on golf performance without a clear political signal.
- `other/unclear`: comments that do not fit the categories above.

The frame coding distinguishes `political`, `blended sport-politics`, `sport`, `depoliticizing bridge`, and `other`.

## Sample Overview

{df_to_markdown(sample_table, index=False)}

## Stance And Frame Distribution

The coded distribution shows a conversation saturated with political reference, but not one dominated by explicit ideological declaration. The largest category is **Trump-referential/unclear** (`{int(stance_table.iloc[0]["comments"]):,}` comments), followed by **other/unclear**. Explicit **pro-Trump/supportive** comments (`{pro_count}`) substantially outnumber explicit **anti-Trump/oppositional** comments (`{anti_count}`), while overt **depoliticizing/bridge** comments are rare (`{bridge_count}`).

On the frame side, **{political_share:.1%}** of comments fall into a political frame and another **{blended_share:.1%}** into a blended sport-politics frame, which indicates that Trump's presence is not being discussed as a purely athletic event.

### Stance Distribution

{df_to_markdown(stance_table, index=False)}

![Stance distribution](../outputs/figures/stance_distribution.png)

### Frame Distribution

{df_to_markdown(frame_table, index=False)}

![Frame distribution](../outputs/figures/frame_distribution.png)

## Engagement Across Stance Positions

The grouped engagement table suggests that the discussion is not symmetrical across stance positions.

- `sport-centered` comments have the highest median comment views (`{sport_views:.1f}`), which suggests that purely golf-centered comments can still travel widely.
- `pro-Trump/supportive` comments have a median of `{pro_likes:.1f}` likes, higher than the median for explicitly anti-Trump comments.
- `anti-Trump/oppositional` comments have low visibility in median terms (`{anti_views:.1f}` median views), even though some individual anti comments still attract attention.

### Engagement By Stance

{df_to_markdown(stance_engagement, index=False)}

![Engagement by stance](../outputs/figures/engagement_by_stance.png)

## High-Engagement Comment Composition

The top-decile comparison shows how stance categories are represented among the most visible comments. Explicitly supportive comments make up **{top_likes_pro:.1%}** of the top-like decile, compared with **{top_likes_anti:.1%}** for explicitly oppositional comments. The top-decile comments are still dominated by the broad `Trump-referential/unclear` category, which suggests that much of the conversation's visible interaction sits in a politically charged but not always explicitly resolved middle zone.

{df_to_markdown(top_decile_table, index=False)}

![Top-decile stance shares](../outputs/figures/top_decile_stance_shares.png)

## Controlled Checks

To preserve continuity with the earlier pipeline, the script still runs simple controlled models for `log(1 + likes)` and `log(1 + Comment views)`. In these models, the stance categories are compared against the `Trump-referential/unclear` reference group while controlling for author's attention quantity, verification status, language, media outlets, and author posting volume.

The most consistent controlled result is that **author's attention quantity remains a strong predictor of visible reaction**, even after the discourse pivot:

- For likes: `b = {like_attention_row["coef"]:.3f}`, `p {"<0.001" if like_attention_row["p_value"] < 0.001 else f"= {like_attention_row['p_value']:.3f}"}`.
- For comment views: `b = {views_attention_row["coef"]:.3f}`, `p {"<0.001" if views_attention_row["p_value"] < 0.001 else f"= {views_attention_row['p_value']:.3f}"}`.

However, the explicit stance coefficients are less stable, largely because the anti-Trump and depoliticizing categories are comparatively small. That makes the current stance modeling useful as an exploratory supplement rather than as the final causal claim.

### Controlled Model Summary

{df_to_markdown(model_summary, index=False)}

## Interpretation

This revised analysis better fits the P3 prompt because it treats the Break 50 comments as **instances of contested discourse**, not just as engagement datapoints. The main substantive finding is that the comment space is strongly politicized: many comments reference Trump directly, and a large share of the conversation frames the event politically or as a blend of politics and golf. At the same time, relatively few comments use explicit depoliticizing language, which suggests that the conversation rarely escapes the politics-sport overlap even when commenters want it to.

The analysis also shows that supportive, oppositional, and depoliticizing comments do not receive identical reaction patterns. The visible conversation is not purely anti or purely pro; rather, it is dominated by a large referential middle category in which commenters engage Trump as a presence, symbol, or controversy without always taking a sharply coded stance.

## Limitations

- The discourse coding is heuristic and dictionary-based rather than hand-coded.
- Some comments are duplicated or formulaic, which can inflate visible categories.
- The `Trump-referential/unclear` category is intentionally broad, so it captures ambiguity at the cost of precision.
- Because the anti-Trump and depoliticizing categories are small, controlled stance estimates should be interpreted cautiously.

## Reproduction

```powershell
$env:BREAK50_DATA_PATH=\"{data_path}\"
py -X utf8 scripts/run_analysis.py
```
"""
    (DOCS_DIR / "analysis_report.md").write_text(report, encoding="utf-8")


def write_summary_json(
    data_path: Path,
    df: pd.DataFrame,
    stance_table: pd.DataFrame,
    frame_table: pd.DataFrame,
    engagement_table: pd.DataFrame,
) -> None:
    payload = {
        "data_path": str(data_path),
        "n_comments": int(len(df)),
        "stance_distribution": stance_table.to_dict(orient="records"),
        "frame_distribution": frame_table.to_dict(orient="records"),
        "engagement_by_stance": engagement_table.to_dict(orient="records"),
    }
    (TABLES_DIR / "analysis_summary.json").write_text(json.dumps(payload, indent=2), encoding="utf-8")


def main() -> None:
    ensure_dirs()
    data_path = resolve_data_path()
    df = load_and_prepare(data_path)

    sample_table = sample_overview(df)
    stats_table = descriptive_stats(df)
    stance_table = stance_distribution(df)
    frame_table = frame_distribution(df)
    stance_engagement = engagement_by_stance(df)
    frame_engagement = engagement_by_frame(df)
    top_decile_table = top_decile_shares(df)
    model_details, model_summary = discourse_model_summary(df)

    write_markdown_table(sample_table, TABLES_DIR / "sample_overview.md", index=False)
    write_markdown_table(stats_table, TABLES_DIR / "descriptive_stats.md", index=True)
    write_markdown_table(stance_table, TABLES_DIR / "stance_distribution.md", index=False)
    write_markdown_table(frame_table, TABLES_DIR / "frame_distribution.md", index=False)
    write_markdown_table(stance_engagement, TABLES_DIR / "engagement_by_stance.md", index=False)
    write_markdown_table(frame_engagement, TABLES_DIR / "engagement_by_frame.md", index=False)
    write_markdown_table(top_decile_table, TABLES_DIR / "top_decile_stance.md", index=False)
    write_markdown_table(model_details, TABLES_DIR / "discourse_model_details.md", index=False)
    write_markdown_table(model_summary, TABLES_DIR / "discourse_model_summary.md", index=False)

    create_stance_distribution_figure(stance_table)
    create_frame_distribution_figure(frame_table)
    create_engagement_by_stance_figure(stance_engagement)
    create_top_decile_figure(top_decile_table)
    build_analysis_report(
        data_path,
        df,
        sample_table,
        stance_table,
        frame_table,
        stance_engagement,
        top_decile_table,
        model_summary,
    )
    write_summary_json(data_path, df, stance_table, frame_table, stance_engagement)
    print(f"Analysis complete. Report written to: {DOCS_DIR / 'analysis_report.md'}")


if __name__ == "__main__":
    main()
