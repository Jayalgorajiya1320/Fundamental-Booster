# 📊 Excel Guide: Retrieve Product Prices Using Product Codes

## 🔹 Overview
This document explains how to retrieve product prices based on product codes using Excel formulas. It includes step-by-step explanations, formulas, and examples.

---

## 🔹 Dataset Structure

| Product Code | Price |
|--------------|------|
| P101         | 500  |
| P102         | 800  |
| P103         | 1200 |

---

## 🔹 Objective
Given a product code, retrieve the corresponding price automatically.

---

## 🔹 Method 1: VLOOKUP

### Formula:
```
=VLOOKUP(E2, A2:B100, 2, FALSE)
```

### Explanation:
- `E2` → Product code to search
- `A2:B100` → Table range
- `2` → Column number (Price column)
- `FALSE` → Exact match

---

## 🔹 Method 2: XLOOKUP (Recommended)

### Formula:
```
=XLOOKUP(E2, A2:A100, B2:B100)
```

### With Error Handling:
```
=XLOOKUP(E2, A2:A100, B2:B100, "Not Found")
```

### Explanation:
- Lookup value: E2
- Lookup array: Product Code column
- Return array: Price column

---

## 🔹 Method 3: INDEX + MATCH

### Formula:
```
=INDEX(B2:B100, MATCH(E2, A2:A100, 0))
```

### Explanation:
- MATCH finds row number
- INDEX returns value from Price column

---

## 🔹 Absolute Reference Usage

### Example:
```
=VLOOKUP(E2, $A$2:$B$100, 2, FALSE)
```

### Why?
- `$` locks the range so it does not change when dragging formula

---

## 🔹 Comparison Table

| Method        | Use Case              | Difficulty |
|--------------|---------------------|-----------|
| VLOOKUP      | Basic lookup         | Easy      |
| XLOOKUP      | Modern & flexible    | Easy      |
| INDEX+MATCH  | Advanced scenarios   | Medium    |

---

## 🔹 Pro Tips
- Always use exact match (`FALSE`)
- Use `$` for large datasets
- Prefer XLOOKUP in newer Excel versions

---

https://drive.google.com/file/d/1Okyk85pIxfZOBbZGJ_RDc1icamIdr2eW/view?usp=drive_link
