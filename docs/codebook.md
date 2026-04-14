# Break 50 Discourse Codebook

This codebook documents the exploratory coding used in `scripts/run_analysis.py`. The goal is not to claim perfect classification, but to make the coding logic transparent and reproducible.

## 1. Stance Categories

### `pro-Trump/supportive`

Comments are placed in this category when they mention Trump and include clearly supportive or celebratory language. Example signals include words or phrases such as:

- `greatest`
- `elite`
- `legendary`
- `awesome`
- `incredible`
- `has game`
- `can play`
- `national treasure`
- `goat`

### `anti-Trump/oppositional`

Comments are placed in this category when they contain condemnation, rejection, or moral critique directed at Trump or Bryson for platforming him. Example signals include:

- `felon`
- `racist`
- `rapist`
- `pedo`
- `cheat`
- `boycott`
- `be better`
- `disgusted`
- `polarizing person`

### `depoliticizing/bridge`

Comments are placed in this category when they explicitly try to suspend political conflict in favor of golf, civility, or shared enjoyment. Example signals include:

- `no politics`
- `politics out of it`
- `love of golf`
- `healing`
- `play together`
- `very little politicin`

### `Trump-referential/unclear`

Comments are placed in this category when they mention Trump or related political markers but do not clearly resolve into supportive, oppositional, or depoliticizing language under the heuristic rules.

### `sport-centered`

Comments are placed in this category when they focus on golf performance or the Break 50 format without a clear political cue.

### `other/unclear`

Comments that do not fit the categories above are coded as `other/unclear`.

## 2. Frame Categories

### `political`

Comments that use political language but do not strongly invoke golf-performance terms.

### `blended sport-politics`

Comments that combine political and golf-performance language.

### `sport`

Comments that discuss golf performance, the Break 50 challenge, or Bryson without meaningful political language.

### `depoliticizing bridge`

Comments that explicitly call for politics to be set aside or reframed through sport.

### `other`

Comments that do not fit the categories above.

## 3. Additional Indicators

### `moral_condemnation`

Binary indicator marking comments that contain explicit condemnation or attack language such as `felon`, `racist`, `cheat`, or `boycott`.

### `depoliticizing_appeal`

Binary indicator marking comments that explicitly argue for golf, civility, or shared experience over politics.

## 4. Important Caveat

This coding scheme is intentionally transparent but imperfect. It should be read as a first-pass discourse mapping strategy, not as a substitute for full manual coding or intercoder reliability testing.
