# How to Use Placeholders

## What Are Placeholders?

Placeholders are **special text** in your Word template that automatically get replaced with customer data from your Excel file.

## How to Create Placeholders

In your Word template, write any placeholder using this format:
```
{COLUMN_NAME}
```

Where `COLUMN_NAME` is **any column name from your Excel file**.

## Example

### Your Excel File has these columns:
| SSA | Billing Account | CUSTOMER NAME | Department | Outstanding amount in Rs |
|-----|-----------------|---------------|------------|---------------------------|
| SSA-001 | ACC-12345 | John Smith | Sales | 5000.50 |
| SSA-002 | ACC-67890 | Jane Doe | Support | 2500.75 |

### In Your Word Template, you can use:

**Exact column names:**
```
{CUSTOMER NAME}
{Outstanding amount in Rs}
{Billing Account}
{Department}
{SSA}
```

**Underscores instead of spaces:**
```
{CUSTOMER_NAME}
{Outstanding_amount_in_Rs}
```

**No spaces (combined):**
```
{CUSTOMERNAME}
{Outstandingamountin Rs}
```

**Lowercase:**
```
{customer name}
{outstanding amount in rs}
```

**Uppercase:**
```
{CUSTOMER NAME}
{OUTSTANDING AMOUNT IN RS}
```

---

## Real Example Template Text

```
Dear {CUSTOMER NAME},

This is to remind you of outstanding dues for account {Billing Account}.

Outstanding Amount: Rs. {Outstanding amount in Rs}/-
Department: {Department}

Please settle the amount at your earliest convenience.

Regards
```

---

## Generated Output (for John Smith)

```
Dear John Smith,

This is to remind you of outstanding dues for account ACC-12345.

Outstanding Amount: Rs. 5000.50/-
Department: Sales

Please settle the amount at your earliest convenience.

Regards
```

---

## Step-by-Step Usage

1. **Create your Word template** with placeholders like `{CUSTOMER NAME}`, `{Outstanding amount in Rs}`, etc.

2. **Save as .docx file** (not PDF or other formats)

3. **Upload in the app:**
   - Upload your Excel file (Step 1)
   - Upload your Word template (Step 2)
   - The app will **detect all placeholders** automatically
   - Click "Generate Letters" (Step 3)

4. **Download the generated documents** - each customer gets a personalized letter!

---

## Placeholder Matching Rules

The app automatically handles:
- ✅ Spaces in names: `{CUSTOMER NAME}` matches column `CUSTOMER NAME`
- ✅ Underscores: `{CUSTOMER_NAME}` matches column `CUSTOMER NAME`
- ✅ No spaces: `{CUSTOMERNAME}` matches column `CUSTOMER NAME`
- ✅ Case insensitive: `{customer name}` matches column `CUSTOMER NAME`

---

## Common Placeholders (based on your Excel structure)

| Placeholder | What It Contains | Example |
|---|---|---|
| `{CUSTOMER NAME}` | Customer name | John Smith |
| `{Billing Account}` | Account number | ACC-12345 |
| `{Outstanding amount in Rs}` | Outstanding amount | 5000.50 |
| `{Department}` | Department | Sales |
| `{Address}` | Customer address | 123 Main St |
| `{SSA}` | SSA code | SSA-001 |
| `{Status(Active/Inactive)}` | Account status | Active |
| `{Accot Subtype}` | Account subtype | Premium |

**Note:** Any column from your Excel file can be used as a placeholder!

---

## Tips

- Always use **curly braces** `{ }` around placeholder names
- Match the **exact column names** from your Excel file
- You can use **multiple variations** in your template (all will work)
- The app will show you **detected placeholders** after you upload the template
