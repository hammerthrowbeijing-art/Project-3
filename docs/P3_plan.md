## 1. Research Questions & Significance

For each research question, address the phenomenon under investigation, the question being asked, and why it matters.

### RQ 1

The Question: This project asks how commenters frame Donald Trump's appearance in the Break 50 discussion on X, and whether the discussion is organized primarily as sport talk, political talk, or a contested blend of both.

The Context: Break 50 is analytically valuable because Trump's appearance on a golf show is not just another sports-media event. It is a politically charged crossover moment in which commenters negotiate whether sport should remain separate from politics, whether Trump's presence is acceptable or objectionable, and whether golf can function as a neutral space at all. This makes the comment section a site of contested discourse rather than a simple pool of engagement metrics.

Why It Matters: This question matters because it speaks directly to how sport and politics become entangled in contemporary platform culture. For researchers, it provides a case of boundary-making and polarization in public discourse. For sports-media producers and athletes, it clarifies how politically polarizing guest appearances reshape audience conversation. For audiences, it reveals whether the visible comment climate is dominated by support, condemnation, depoliticizing appeals, or unresolved political reference.

### RQ 2

The Question: How are stance positions distributed across the Break 50 comments, specifically pro-Trump/supportive, anti-Trump/oppositional, depoliticizing/bridge, Trump-referential/unclear, sport-centered, and other/unclear comments?

The Context: The core P3 issue is not simply whether comments get attention, but how political disagreement is expressed and structured. A stance distribution helps identify whether the conversation is polarized into explicit camps, dominated by moral condemnation, softened by depoliticizing appeals, or organized around more ambiguous references.

Why It Matters: This matters because the difference between a polarized comment field and an ambiguous-but-politicized one changes how we interpret platform discourse. A conversation full of explicit support and opposition looks different from one in which politics is present but often only indirectly resolved.

### RQ 3

The Question: Do visible reaction metrics such as likes, comment views, and retweet presence differ across stance positions and discourse frames?

The Context: Once stance and frame are identified, the next question is whether some discourse positions receive more visible reinforcement than others. This addresses whether platform reaction clusters around certain forms of political expression, such as supportive comments, oppositional comments, or sport-centered comments that avoid explicit political judgment.

Why It Matters: This matters because reaction patterns help show which discursive positions become most visible and potentially most influential. It also creates a better use for engagement metrics than treating them as interchangeable dependent variables.

## 2. Dataset Selection & Justification

Dataset Choice: Break 50 X/Twitter comment dataset

Justification: This project uses only `B50_X_COMMENT.xlsx`, which contains 1,008 public comments collected between July 23, 2024 and August 1, 2024. The file is appropriate because it contains the full comment text, engagement variables, author metadata, and timing information needed to examine discourse structure within a single event-centered discussion. The project intentionally stays within one dataset so that stance coding, frame coding, and reaction measures are all drawn from the same discursive environment.

Key file used locally: `B50_X_COMMENT.xlsx`

## 3. Preliminary Variable Operationalization

| Construct | Operational Definition | Data Source / Indicator |
| --- | --- | --- |
| Stance toward Trump's appearance | Comments will be coded into pro-Trump/supportive, anti-Trump/oppositional, depoliticizing/bridge, Trump-referential/unclear, sport-centered, or other/unclear categories using transparent lexical rules and close reading logic. | Comment text in `contents` |
| Discourse frame | Comments will be coded as political, sport, blended sport-politics, depoliticizing bridge, or other depending on whether the language emphasizes political identity, golf performance, both, or explicit attempts to suspend politics. | Comment text in `contents` |
| Moral condemnation / boundary policing | Comments containing condemnation language such as `felon`, `racist`, `cheat`, `boycott`, or related evaluative signals will be marked as moralized opposition. | Comment text in `contents` |
| Depoliticizing appeal | Comments explicitly calling for politics to be set aside in favor of golf, shared enjoyment, or civility will be marked as depoliticizing appeals. | Comment text in `contents` |
| Reaction metrics | Likes, retweets, and comment views will be used as secondary outcomes to evaluate which discourse positions receive more visible reinforcement. | `likes`, `retweets count`, `Comment views` |
| Author-side controls | Author's attention quantity, verification status, language, number of media outlets, and number of author posts will be used as controls rather than as the primary theoretical focus. | `Author's attention quantity`, `blue_verified`, `Comment language`, `number of media outlets`, `number of author posts` |

## 4. Proposed Analyses

| Analysis Type | Description | RQ Addressed |
| --- | --- | --- |
| Descriptive profiling | Summarize the sample, date range, comment counts, and skew in reaction metrics. | RQ 1, RQ 3 |
| Stance distribution analysis | Count how many comments fall into each stance category and evaluate whether the conversation is dominated by explicit support, explicit opposition, ambiguity, or depoliticizing appeals. | RQ 2 |
| Frame analysis | Measure how often comments frame the event as political, sport-only, or blended sport-politics. | RQ 1 |
| Engagement-by-stance comparison | Compare likes, views, and retweet rates across stance categories. | RQ 3 |
| Top-decile composition check | Compare the stance mix in the full dataset to the stance mix among the highest-like and highest-view comments. | RQ 3 |
| Controlled exploratory models | Model log-likes and log-views using stance categories while controlling for author's attention quantity and author metadata. | RQ 3 |

## 5. Limitations & Potential Issues

- The stance and frame coding is heuristic and keyword-based rather than hand-coded by multiple coders.
- Some comments are duplicated or formulaic, which may inflate certain categories.
- The anti-Trump and depoliticizing categories are relatively small, which limits model stability.
- The dataset covers one highly specific event and one short time window, so generalizability is limited.
- The data are observational and public-platform based, so the study identifies visible discourse patterns rather than causal effects.

## 6. Ethical Considerations

This project uses public comments rather than private communications, so privacy risks are relatively low. Even so, the analysis should avoid unnecessary exposure of usernames and should report results in aggregate whenever possible. Because the dataset contains political conflict, moral condemnation, and personal attack language, the analysis should avoid re-amplifying harmful statements beyond what is needed for scholarly interpretation. Bias can enter through platform algorithms, the lexical coding rules, and the fact that the dataset captures only a single collection window.

## 7. Group Role Assignments

| Role | Group Member | Primary Responsibilities |
| --- | --- | --- |
| Project & Data Lead | Larry Nie | Dataset inspection, coding design, and file management |
| Methods Lead | Larry Nie | Discourse operationalization, exploratory modeling, and result interpretation |
| Writing & Visualization Lead | Larry Nie | Report drafting, markdown documentation, and figure production |

## 8. Data Visualization Plan

Primary Goal: The main figures will show how the comment field is distributed across stance and frame categories, and whether highly visible comments are disproportionately associated with particular discourse positions.

Visualization Description:

- A stance distribution chart showing the number and share of comments in each stance category
- A frame distribution chart showing how often the event is discussed as political, sport-centered, or blended
- An engagement-by-stance chart comparing median likes and median views across stance positions
- A top-decile composition figure comparing overall stance shares with the stance shares among the most-liked and most-viewed comments

Design Rationale: These visuals are better aligned with the P3 objective because they make disagreement, ambiguity, and depoliticizing language visible. They treat engagement as a secondary outcome layered onto discourse positions rather than as the study's only subject.

Verification Methods:

- [x] Spot-checked coding outputs against source comments
- [x] Checked that counts sum to the total number of comments
- [x] Verified that figures match the markdown summary tables
- [ ] Had a second coder review classifications

## 9. AI-Assisted Work Documentation & Verification

Tools Used: Codex was used to inspect the dataset, revise the repo structure, script a reproducible analysis pipeline, and help draft the updated project framing after instructor feedback.

Verification Methods:

- [x] Reviewed the code logic for the stance and frame coding rules
- [x] Checked row counts, category counts, and summary statistics against the dataset
- [x] Reviewed the generated interpretations to ensure they match the observed outputs
- [x] Used AI as drafting and scripting support rather than as an unverified source of substantive claims
