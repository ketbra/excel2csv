# Bug: "General" format outputs "Gnral" instead of formatted number

## Description

When using `NumberFormat::parse("General")` and calling `.format()` with a numeric value, the output is the literal text "Gnral" instead of the formatted number.

## Steps to Reproduce

```rust
use ssfmt::{NumberFormat, FormatOptions};

fn main() {
    let opts = FormatOptions::default();
    let fmt = NumberFormat::parse("General").unwrap();

    // Bug: outputs "Gnral" instead of "10"
    println!("{}", fmt.format(10.0, &opts));
}
```

## Expected Behavior

```
10
```

## Actual Behavior

```
Gnral
```

## Root Cause Analysis

The bug has two contributing factors:

### 1. Lexer treats 'e' as ExponentLower token

In `lexer.rs:111-114`:
```rust
'e' if !self.in_bracket => {
    self.advance();
    Token::ExponentLower
}
```

When parsing "General", the two 'e' characters are tokenized as `ExponentLower` instead of `Literal('e')`.

Tokenization of "General":
- `G` → `Literal('G')`
- `e` → `ExponentLower` ← **problem**
- `n` → `Literal('n')`
- `e` → `ExponentLower` ← **problem**
- `r` → `Literal('r')`
- `a` → `Literal('a')`
- `l` → `Literal('l')`

### 2. Scientific parts are not output in literal-only formats

In `formatter/number.rs:130-151`, when there are no digit placeholders, only `FormatPart::Literal` parts are output. The `FormatPart::Scientific` parts created from the 'e' tokens are silently dropped.

### 3. "General" should be special-cased

In Excel, "General" is a built-in format that should use general number formatting (the fallback_format function), not be parsed as literal characters.

## Suggested Fix

Add special handling for "General" format code:

**Option A** (Recommended): In `parser::parse()`, detect "General" (case-insensitive) and return a special empty/default format that triggers fallback formatting:

```rust
pub fn parse(format_code: &str) -> Result<NumberFormat, ParseError> {
    if format_code.is_empty() {
        return Err(ParseError::EmptyFormat);
    }

    // Handle built-in "General" format
    if format_code.eq_ignore_ascii_case("General") {
        return Ok(NumberFormat::from_sections(vec![]));
    }

    // ... rest of parsing
}
```

Then in `NumberFormat::try_format()`, handle empty sections:
```rust
if self.sections.is_empty() {
    return Ok(fallback_format(value));
}
```

**Option B**: In the lexer, only treat 'e'/'E' as exponent tokens when followed by +/- (which is the actual scientific notation pattern like `0.00E+00`).

**Option C**: Recognize "General" as a keyword in the lexer and emit a special token.

## Environment

- ssfmt version: 0.1.0
- Rust version: 1.92.0

## Impact

This affects any Excel cells with the default "General" format that contain numeric values. The formatted output shows "Gnral" instead of the actual number, making the CSV output unusable for those cells.

## Additional Context

- `format_text("10", &opts)` works correctly and returns `"10"` because it bypasses the number formatting path
- Only the numeric `.format()` path is affected
- Discovered while using ssfmt in excel2csv for Excel-to-CSV conversion
