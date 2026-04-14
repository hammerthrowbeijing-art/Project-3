# Break 50 Final Report

## Title

**Contested Discourse in the Break 50 Comment Field: Stance, Framing, and Visible Reaction**

## Abstract

This report examines the Break 50 X/Twitter comment field as a case of contested sport-political discourse rather than as a generic engagement dataset. The core question is how commenters frame Donald Trump's appearance in Break 50, how stance positions are distributed across the discussion, and which discourse positions receive more visible reaction. Using 1,008 public comments collected between July 23, 2024 and August 1, 2024, the analysis applies a transparent lexical coding strategy to classify stance, frame, moral condemnation, and depoliticizing appeals. The findings show that the discussion is strongly politicized: 50.0% of comments are coded as political and another 25.6% as blended sport-politics. The largest stance category is Trump-referential but unresolved, while explicitly supportive comments outnumber explicitly oppositional ones. Visible reaction does not cluster evenly across categories. Sport-centered comments receive the highest median views, explicitly supportive comments receive somewhat stronger median likes than explicitly oppositional comments, and depoliticizing appeals are rare. Overall, the Break 50 comments are best understood as a politically saturated and unevenly negotiated discourse space in which sport and politics are repeatedly made to overlap.

## 1. Introduction

The Break 50 dataset becomes analytically important because it captures more than fan reaction to a golf video. Donald Trump's appearance on a Bryson DeChambeau golf show turns the event into a politically charged media moment. In that setting, the comment section becomes a place where users do not simply react to the content; they negotiate the legitimacy of the guest, the meaning of the appearance, the relation between sport and politics, and the acceptability of other users' reactions.

This matters for P3 because the assignment is not just about correlation or platform metrics. It is about contested discourse. A study that only asks whether author metadata predicts likes or views misses the central feature of the case: the comment field contains disagreement, moral framing, depoliticizing appeals, and boundary-policing around who belongs in a sports-media space and under what terms.

The revised project therefore treats the comments themselves as the object of analysis. The key question is not only which comments get seen, but what kinds of political and sport meanings are being produced.

## 2. Research Questions

This report addresses three linked questions:

1. How is Trump's appearance in Break 50 framed in the comments: as politics, as sport, or as a blend of both?
2. How are stance positions distributed across the comment field, especially supportive, oppositional, depoliticizing, ambiguous, and sport-centered comments?
3. Do visible reaction metrics such as likes and comment views cluster differently across those discourse positions?

In this revised framing, **author's attention quantity** is retained as a secondary control variable rather than the main theoretical focus.

## 3. Data

The analysis uses one local dataset, `B50_X_COMMENT.xlsx`, which is kept outside the public repository in accordance with course policy.

### Sample Characteristics

- Total comments: **1,008**
- Unique usernames: **740**
- Unique comment IDs: **845**
- Source posts: **4**
- Date range: **2024-07-23 02:02:13** to **2024-08-01 02:43:46**
- English-language comments: **920**
- Verified accounts: **833**
- Comments with zero retweets: **908**
- Comments mentioning Trump-related terms under the coding rules: **770**
- Comments with explicit moral-condemnation language: **30**
- Comments with depoliticizing appeals: **9**

These basic counts already suggest a strongly politicized discussion. Trump-related reference appears in most of the dataset, while comments explicitly asking for politics to be suspended are comparatively rare.

## 4. Method

## 4.1 Coding Strategy

Because the dataset does not include pre-coded discourse variables, the analysis uses a transparent lexical coding scheme documented in [codebook.md](./codebook.md).

### Stance categories

- `pro-Trump/supportive`
- `anti-Trump/oppositional`
- `depoliticizing/bridge`
- `Trump-referential/unclear`
- `sport-centered`
- `other/unclear`

### Frame categories

- `political`
- `blended sport-politics`
- `sport`
- `depoliticizing bridge`
- `other`

Two additional indicators were also coded:

- `moral_condemnation`
- `depoliticizing_appeal`

The purpose of this coding is to create a first-pass map of conflict, ambiguity, and bridge language in the Break 50 comments. It is exploratory and should not be treated as a substitute for full manual coding.

## 4.2 Reaction Measures

To examine which discourse positions receive more visible reinforcement, the report uses:

- `likes`
- `Comment views`
- `retweets count` as a secondary descriptive measure

Unlike the earlier version of the project, these are not treated as interchangeable outcomes. Instead, they are used selectively to ask whether certain stance positions are more visible or more endorsed.

## 4.3 Controlled Checks

The analysis also preserves a limited version of the earlier author-side modeling by using `author's attention quantity` as a control variable in exploratory models predicting log-likes and log-views. This allows the report to distinguish discourse-position patterns from simple account-level attention advantages.

## 5. Results

## 5.1 Frame Distribution

The frame results show that the comment field is not primarily organized as neutral sport talk.

| Frame | Comments | Share |
| --- | --- | --- |
| political | 504 | 0.500 |
| blended sport-politics | 258 | 0.256 |
| sport | 41 | 0.041 |
| depoliticizing bridge | 9 | 0.009 |
| other | 196 | 0.194 |

These counts show that fully half of the comments frame the event politically, while another quarter combine political and golf language. Only a small minority are coded as purely sport-centered. This means that Trump's appearance is being processed by the audience as a political event as much as, or more than, a golf-media event.

![Frame distribution](../outputs/figures/frame_distribution.png)

## 5.2 Stance Distribution

The stance distribution adds more nuance. The largest category is not explicit support or opposition, but a broad middle zone of politically charged reference without a sharply coded resolution.

| Stance | Comments | Share |
| --- | --- | --- |
| Trump-referential/unclear | 623 | 0.618 |
| pro-Trump/supportive | 108 | 0.107 |
| anti-Trump/oppositional | 30 | 0.030 |
| depoliticizing/bridge | 9 | 0.009 |
| sport-centered | 41 | 0.041 |
| other/unclear | 197 | 0.195 |

Three points stand out.

1. Explicitly supportive comments outnumber explicitly oppositional comments by more than three to one.
2. Depoliticizing appeals are rare.
3. The dominant category is the Trump-referential middle, which suggests that the comment field is heavily politicized even when users are not always clearly taking sides in a way the heuristic code can sharply classify.

![Stance distribution](../outputs/figures/stance_distribution.png)

## 5.3 Visible Reaction By Stance

The engagement comparison shows that reaction patterns differ across discourse positions.

| Stance | Comments | Median likes | Median views | Retweet rate | Median attention quantity |
| --- | --- | --- | --- | --- | --- |
| Trump-referential/unclear | 623 | 1.0 | 210.0 | 0.098 | 606.0 |
| pro-Trump/supportive | 108 | 2.0 | 234.0 | 0.111 | 565.0 |
| anti-Trump/oppositional | 30 | 1.0 | 82.0 | 0.200 | 346.0 |
| depoliticizing/bridge | 9 | 1.0 | 90.0 | 0.222 | 597.0 |
| sport-centered | 41 | 2.0 | 915.0 | 0.122 | 641.0 |
| other/unclear | 197 | 2.0 | 638.0 | 0.071 | 694.0 |

The clearest pattern is that `sport-centered` comments receive the highest median views. This suggests that comments focused on golf performance, rather than direct political argument, can travel especially widely in the visible conversation. Explicitly supportive comments have somewhat stronger median likes than explicitly oppositional comments. Anti-Trump comments remain present, but their median visibility is much lower.

![Engagement by stance](../outputs/figures/engagement_by_stance.png)

## 5.4 High-Engagement Comment Composition

To see whether highly visible comments differ from the overall discussion, the analysis compares overall stance shares with the top decile of likes and the top decile of comment views.

| Stance | Overall share | Top likes share | Top views share |
| --- | --- | --- | --- |
| Trump-referential/unclear | 0.618 | 0.642 | 0.624 |
| pro-Trump/supportive | 0.107 | 0.123 | 0.089 |
| anti-Trump/oppositional | 0.030 | 0.057 | 0.050 |
| depoliticizing/bridge | 0.009 | 0.019 | 0.000 |
| sport-centered | 0.041 | 0.066 | 0.079 |
| other/unclear | 0.195 | 0.094 | 0.158 |

This comparison suggests that the most visible comments are still dominated by the same broad referential middle that dominates the conversation overall. However, the top-like decile contains somewhat larger shares of both supportive and oppositional comments than their overall baseline, which implies that explicit contestation is not numerically dominant but can still capture attention.

![Top-decile stance shares](../outputs/figures/top_decile_stance_shares.png)

## 5.5 Controlled Checks

The controlled models are not the centerpiece of the revised report, but they provide an important check. After controlling for verification status, language, number of media outlets, number of author posts, and `author's attention quantity`, the attention variable remains strongly associated with visible reaction:

- For `log(1 + likes)`: `b = 0.137`, `p = 0.001`
- For `log(1 + comment views)`: `b = 0.239`, `p < 0.001`

By contrast, the explicit stance coefficients are less stable. This is especially true for `anti-Trump/oppositional` and `depoliticizing/bridge` comments, which are relatively rare in the sample. The controlled models therefore support a narrower conclusion: account-level attention still matters, but it does not replace the need to analyze the discourse categories themselves.

## 6. Discussion

The main contribution of this revised analysis is conceptual. It moves the project into genuine P3 territory by treating the Break 50 comments as a field of **contested discourse** rather than as a set of comments waiting to be correlated with author metadata.

Three broader conclusions follow.

First, the discussion is deeply politicized. A majority of comments are coded as political or blended sport-politics, and only a very small minority are coded as depoliticizing bridge language. This means that Trump's appearance does not remain safely contained within a sports-only frame.

Second, explicit support is more common than explicit opposition in the coded data, but neither side fully dominates the discussion. Instead, the largest space is occupied by Trump-referential comments that keep politics active without always making stance completely explicit. That ambiguity is important: it suggests a discourse field structured less by perfectly separated camps than by overlapping layers of endorsement, condemnation, irony, and unresolved signaling.

Third, visible reaction does not map neatly onto the most openly partisan comments. The most-viewed median pattern belongs to sport-centered comments, while the top-engagement deciles still contain a large share of politically referential but unresolved comments. In other words, the most visible comments are not always the most ideologically explicit ones.

## 7. Limitations

- The coding is heuristic and dictionary-based rather than manually validated by multiple coders.
- Some comments are duplicated or highly formulaic, which may inflate certain categories.
- The `Trump-referential/unclear` category is intentionally broad, so it captures ambiguity at the cost of precision.
- The anti-Trump and depoliticizing categories are small, which reduces statistical stability.
- The dataset covers one event, one platform, and a short time window, so the findings should not be generalized too far.

## 8. Conclusion

The Break 50 comment field is best understood as a politically saturated discourse space in which golf and politics are repeatedly forced together. The dominant pattern is not simply support or opposition, but a large and noisy zone of politically referential commentary. Explicitly supportive comments are more common than explicitly oppositional ones, depoliticizing comments are rare, and the most visible comments are not always the most openly partisan.

This revised framing is closer to the purpose of P3 because it analyzes the comment field as a site where audiences negotiate the boundaries between sport, politics, legitimacy, and visibility. It keeps the useful infrastructure of the earlier repository, but redirects the analytic center toward the contested discourse that makes Break 50 worth studying.

## 9. Repository Notes

This repository now includes:

- a revised P3 plan centered on contested discourse
- a transparent discourse codebook
- a reproducible analysis pipeline
- updated figures and markdown tables
- repo-ready analysis and final reports
- a supplementary appendix preserving the earlier author-attention exploratory analysis

The raw dataset remains intentionally excluded from version control.
