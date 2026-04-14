# Break 50 Analysis Report

## Overview

This revised analysis responds to the instructor feedback by shifting the project from a generic engagement-correlation study to a **P3-style contested-discourse analysis**. Instead of treating all outcomes as interchangeable engagement measures, the report asks how commenters frame Trump's appearance in Break 50, how stance positions are distributed, and which kinds of comments receive more visible reaction.

The raw Excel file remains local and outside version control. Data source used locally: `C:\Users\Larry.Nie\Downloads\B50_X_COMMENT.xlsx`

## Revised Research Question

The revised study asks:

1. How is Trump's appearance in Break 50 framed in the comments: as politics, as sport, or as a blending of the two?
2. How are stance positions distributed across the discussion, specifically pro-Trump/supportive, anti-Trump/oppositional, depoliticizing/bridge, Trump-referential/unclear, sport-centered, and other/unclear comments?
3. Do visible reaction metrics such as likes and comment views cluster differently across those stance positions?

In this revised design, **author's attention quantity** is retained as a secondary control variable rather than the main research question.

## Coding Strategy

Because the dataset does not include pre-coded discourse variables, the pipeline applies a transparent keyword-based coding scheme to approximate stance and frame. The coding rules are documented in [codebook.md](./codebook.md).

- `pro-Trump/supportive`: comments using supportive or celebratory language toward Trump or his appearance.
- `anti-Trump/oppositional`: comments using condemnation, moral critique, or rejection language.
- `depoliticizing/bridge`: comments explicitly calling for politics to be set aside in favor of golf or shared enjoyment.
- `Trump-referential/unclear`: comments that mention Trump but do not clearly resolve into support or opposition under the heuristic rules.
- `sport-centered`: comments focused on golf performance without a clear political signal.
- `other/unclear`: comments that do not fit the categories above.

The frame coding distinguishes `political`, `blended sport-politics`, `sport`, `depoliticizing bridge`, and `other`.

## Sample Overview

| metric | value |
| --- | --- |
| Comments | 1,008 |
| Unique usernames | 740 |
| Unique comment IDs | 845 |
| Unique source posts | 4 |
| Date start | 2024-07-23 02:02:13 |
| Date end | 2024-08-01 02:43:46 |
| English-language comments | 920 |
| Verified accounts | 833 |
| Comments with zero retweets | 908 |
| Comments mentioning Trump-related terms | 770 |
| Comments with moral-condemnation language | 30 |
| Comments with depoliticizing appeals | 9 |

## Stance And Frame Distribution

The coded distribution shows a conversation saturated with political reference, but not one dominated by explicit ideological declaration. The largest category is **Trump-referential/unclear** (`623` comments), followed by **other/unclear**. Explicit **pro-Trump/supportive** comments (`108`) substantially outnumber explicit **anti-Trump/oppositional** comments (`30`), while overt **depoliticizing/bridge** comments are rare (`9`).

On the frame side, **50.0%** of comments fall into a political frame and another **25.6%** into a blended sport-politics frame, which indicates that Trump's presence is not being discussed as a purely athletic event.

### Stance Distribution

| stance | comments | share |
| --- | --- | --- |
| Trump-referential/unclear | 623 | 0.618 |
| pro-Trump/supportive | 108 | 0.107 |
| anti-Trump/oppositional | 30 | 0.030 |
| depoliticizing/bridge | 9 | 0.009 |
| sport-centered | 41 | 0.041 |
| other/unclear | 197 | 0.195 |

![Stance distribution](../outputs/figures/stance_distribution.png)

### Frame Distribution

| frame | comments | share |
| --- | --- | --- |
| political | 504 | 0.500 |
| blended sport-politics | 258 | 0.256 |
| sport | 41 | 0.041 |
| depoliticizing bridge | 9 | 0.009 |
| other | 196 | 0.194 |

![Frame distribution](../outputs/figures/frame_distribution.png)

## Engagement Across Stance Positions

The grouped engagement table suggests that the discussion is not symmetrical across stance positions.

- `sport-centered` comments have the highest median comment views (`915.0`), which suggests that purely golf-centered comments can still travel widely.
- `pro-Trump/supportive` comments have a median of `2.0` likes, higher than the median for explicitly anti-Trump comments.
- `anti-Trump/oppositional` comments have low visibility in median terms (`82.0` median views), even though some individual anti comments still attract attention.

### Engagement By Stance

| stance | comments | likes_mean | likes_median | views_mean | views_median | retweet_rate | attention_median |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Trump-referential/unclear | 623 | 37.881 | 1.000 | 7,809.79 | 210.000 | 0.098 | 606.000 |
| pro-Trump/supportive | 108 | 36.602 | 2.000 | 7,684.05 | 234.000 | 0.111 | 565.000 |
| anti-Trump/oppositional | 30 | 10.400 | 1.000 | 4,659.67 | 82.000 | 0.200 | 346.000 |
| depoliticizing/bridge | 9 | 17.556 | 1.000 | 1,691.67 | 90.000 | 0.222 | 597.000 |
| sport-centered | 41 | 42.634 | 2.000 | 10,931.29 | 915.000 | 0.122 | 641.000 |
| other/unclear | 197 | 16.036 | 2.000 | 3,601.58 | 638.000 | 0.071 | 694.000 |

![Engagement by stance](../outputs/figures/engagement_by_stance.png)

## High-Engagement Comment Composition

The top-decile comparison shows how stance categories are represented among the most visible comments. Explicitly supportive comments make up **12.3%** of the top-like decile, compared with **5.7%** for explicitly oppositional comments. The top-decile comments are still dominated by the broad `Trump-referential/unclear` category, which suggests that much of the conversation's visible interaction sits in a politically charged but not always explicitly resolved middle zone.

| stance | overall_share | top_likes_share | top_views_share |
| --- | --- | --- | --- |
| Trump-referential/unclear | 0.618 | 0.642 | 0.624 |
| pro-Trump/supportive | 0.107 | 0.123 | 0.089 |
| anti-Trump/oppositional | 0.030 | 0.057 | 0.050 |
| depoliticizing/bridge | 0.009 | 0.019 | 0.000 |
| sport-centered | 0.041 | 0.066 | 0.079 |
| other/unclear | 0.195 | 0.094 | 0.158 |

![Top-decile stance shares](../outputs/figures/top_decile_stance_shares.png)

## Controlled Checks

To preserve continuity with the earlier pipeline, the script still runs simple controlled models for `log(1 + likes)` and `log(1 + Comment views)`. In these models, the stance categories are compared against the `Trump-referential/unclear` reference group while controlling for author's attention quantity, verification status, language, media outlets, and author posting volume.

The most consistent controlled result is that **author's attention quantity remains a strong predictor of visible reaction**, even after the discourse pivot:

- For likes: `b = 0.137`, `p <0.001`.
- For comment views: `b = 0.239`, `p <0.001`.

However, the explicit stance coefficients are less stable, largely because the anti-Trump and depoliticizing categories are comparatively small. That makes the current stance modeling useful as an exploratory supplement rather than as the final causal claim.

### Controlled Model Summary

| outcome | comparison_to_reference | coef | robust_se | ci_low | ci_high | p_value | r2 | n |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Likes (log1p) | pro-Trump/supportive vs Trump-referential/unclear | 0.171 | 0.160 | -0.142 | 0.485 | 0.284 | 0.090 | 1,008 |
| Likes (log1p) | anti-Trump/oppositional vs Trump-referential/unclear | 0.381 | 0.319 | -0.244 | 1.007 | 0.232 | 0.090 | 1,008 |
| Likes (log1p) | depoliticizing/bridge vs Trump-referential/unclear | 0.169 | 0.541 | -0.892 | 1.229 | 0.755 | 0.090 | 1,008 |
| Likes (log1p) | sport-centered vs Trump-referential/unclear | 0.427 | 0.264 | -0.091 | 0.945 | 0.106 | 0.090 | 1,008 |
| Likes (log1p) | other/unclear vs Trump-referential/unclear | -0.160 | 0.102 | -0.360 | 0.040 | 0.117 | 0.090 | 1,008 |
| Likes (log1p) | log_attention control effect | 0.137 | 0.041 | 0.056 | 0.218 | 0.001 | 0.090 | 1,008 |
| Comment views (log1p) | pro-Trump/supportive vs Trump-referential/unclear | 0.177 | 0.239 | -0.292 | 0.645 | 0.459 | 0.129 | 1,008 |
| Comment views (log1p) | anti-Trump/oppositional vs Trump-referential/unclear | 0.667 | 0.571 | -0.452 | 1.785 | 0.243 | 0.129 | 1,008 |
| Comment views (log1p) | depoliticizing/bridge vs Trump-referential/unclear | -0.310 | 0.793 | -1.864 | 1.243 | 0.695 | 0.129 | 1,008 |
| Comment views (log1p) | sport-centered vs Trump-referential/unclear | 0.742 | 0.377 | 0.003 | 1.482 | 0.049 | 0.129 | 1,008 |
| Comment views (log1p) | other/unclear vs Trump-referential/unclear | 0.673 | 0.168 | 0.345 | 1.002 | 0.000 | 0.129 | 1,008 |
| Comment views (log1p) | log_attention control effect | 0.239 | 0.063 | 0.116 | 0.363 | 0.000 | 0.129 | 1,008 |

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
$env:BREAK50_DATA_PATH="C:\Users\Larry.Nie\Downloads\B50_X_COMMENT.xlsx"
py -X utf8 scripts/run_analysis.py
```
