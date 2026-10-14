# Building a CRM with RexxJS Spreadsheet

This tutorial demonstrates how to build a Customer Relationship Management (CRM) system using the advanced features of RexxJS Spreadsheet. You'll learn how to use query chaining, table metadata, context-aware validation, auto-increment IDs, and batch operations to create a functional CRM.

## What You'll Build

By the end of this tutorial, you'll have a CRM spreadsheet that can:
- ✅ Track customers with auto-generated IDs
- ✅ Manage orders with automatic order numbers
- ✅ Query sales data by region, product, and customer
- ✅ Validate data entry to prevent duplicates and errors
- ✅ Generate sales reports and analytics
- ✅ Bulk import data from external sources

## Prerequisites

- RexxJS Spreadsheet installed (web or desktop mode)
- Basic understanding of spreadsheet formulas
- Familiarity with JavaScript/RexxJS syntax

---

## Step 1: Set Up the Customer Table

We'll start by creating a customer database with auto-incrementing customer IDs.

### 1.1 Create Column Headers

In a new spreadsheet, add the following headers in row 1:

| A | B | C | D | E |
|---|---|---|---|---|
| **CustomerID** | **CompanyName** | **ContactName** | **Email** | **Region** |

### 1.2 Configure Auto-Increment IDs

Use the Control Bus to configure automatic customer ID generation:

```javascript
// Via Control Bus (desktop mode with HTTP API)
CALL CONFIGURE_AUTO_ID("A", 1000, "CUST-")

// IDs will be: CUST-1000, CUST-1001, CUST-1002, etc.
```

### 1.3 Add Data Validation

Set up validation rules to ensure data quality:

```javascript
// Email must be unique (only check on creation)
model.setCellValidation('D2:D1000', {
  type: 'contextual',
  onCreate: 'UNIQUE("D:D", value)',
  onUpdate: 'true'  // Allow updates
});

// Email format validation
model.setCellValidation('D2:D1000', {
  type: 'formula',
  formula: 'value CONTAINS "@"'
});

// Region must be one of: North, South, East, West
model.setCellValidation('E2:E1000', {
  type: 'list',
  values: ['North', 'South', 'East', 'West']
});
```

### 1.4 Define Table Metadata

Create a table definition for easy querying:

```javascript
CALL SET_TABLE_METADATA("Customers", '{
  "range": "A1:E1000",
  "columns": {
    "customer_id": "A",
    "company": "B",
    "contact": "C",
    "email": "D",
    "region": "E"
  },
  "hasHeader": true,
  "types": {
    "customer_id": "string",
    "company": "string",
    "contact": "string",
    "email": "string",
    "region": "string"
  },
  "descriptions": {
    "customer_id": "Unique customer identifier (auto-generated)",
    "company": "Company or organization name",
    "contact": "Primary contact person",
    "email": "Contact email address (must be unique)",
    "region": "Sales region (North, South, East, West)"
  }
}')
```

### 1.5 Add Sample Customers

Now insert some rows. The customer IDs will be automatically generated:

```javascript
// Using batch operations for efficiency
CALL BATCH_SET_CELLS('[
  {"address": "B2", "value": "Acme Corp"},
  {"address": "C2", "value": "John Smith"},
  {"address": "D2", "value": "john@acme.com"},
  {"address": "E2", "value": "West"},

  {"address": "B3", "value": "TechStart Inc"},
  {"address": "C3", "value": "Jane Doe"},
  {"address": "D3", "value": "jane@techstart.com"},
  {"address": "E3", "value": "North"},

  {"address": "B4", "value": "Global Solutions"},
  {"address": "C4", "value": "Bob Johnson"},
  {"address": "D4", "value": "bob@global.com"},
  {"address": "E4", "value": "East"}
]')
```

After inserting rows, column A will automatically show: `CUST-1000`, `CUST-1001`, `CUST-1002`

---

## Step 2: Create the Orders Table

Now we'll create an orders table that references customers.

### 2.1 Create Headers (Sheet2 or columns F-K)

For this tutorial, let's create a second sheet called "Orders":

| A | B | C | D | E | F |
|---|---|---|---|---|---|
| **OrderID** | **CustomerID** | **OrderDate** | **Product** | **Quantity** | **Amount** |

### 2.2 Configure Order Auto-IDs

```javascript
// Switch to Orders sheet first
model.setActiveSheet('Orders');

// Configure auto-IDs for orders
CALL CONFIGURE_AUTO_ID("A", 5000, "ORD-")
```

### 2.3 Add Validation for Customer References

```javascript
// Customer ID must exist in Customers table
model.setCellValidation('B2:B1000', {
  type: 'formula',
  formula: 'FIND_ROW_BY_ID(value) != ""'
});

// Quantity must be positive
model.setCellValidation('E2:E1000', {
  type: 'formula',
  formula: 'value > 0'
});

// Amount must be positive
model.setCellValidation('F2:F1000', {
  type: 'formula',
  formula: 'value > 0'
});
```

### 2.4 Define Orders Table Metadata

```javascript
CALL SET_TABLE_METADATA("Orders", '{
  "range": "A1:F5000",
  "columns": {
    "order_id": "A",
    "customer_id": "B",
    "order_date": "C",
    "product": "D",
    "quantity": "E",
    "amount": "F"
  },
  "hasHeader": true,
  "types": {
    "order_id": "string",
    "customer_id": "string",
    "order_date": "string",
    "product": "string",
    "quantity": "number",
    "amount": "number"
  }
}')
```

### 2.5 Add Sample Orders

```javascript
CALL BATCH_SET_CELLS('[
  {"address": "B2", "value": "CUST-1000"},
  {"address": "C2", "value": "2025-01-15"},
  {"address": "D2", "value": "Widget Pro"},
  {"address": "E2", "value": "10"},
  {"address": "F2", "value": "1500"},

  {"address": "B3", "value": "CUST-1001"},
  {"address": "C3", "value": "2025-01-16"},
  {"address": "D3", "value": "Gadget Plus"},
  {"address": "E3", "value": "5"},
  {"address": "F3", "value": "850"},

  {"address": "B4", "value": "CUST-1000"},
  {"address": "C4", "value": "2025-01-20"},
  {"address": "D4", "value": "Super Gizmo"},
  {"address": "E4", "value": "15"},
  {"address": "F4", "value": "3200"}
]')
```

---

## Step 3: Build Analytics with Query Chaining

Now let's create powerful analytics using the TABLE() function and query chaining.

### 3.1 Total Sales by Customer

Create a new sheet called "Analytics" and add this formula in cell B2:

```javascript
=TABLE('Orders') |> GROUP_BY('customer_id') |> SUM('amount')
```

This returns an object like:
```javascript
{
  "CUST-1000": 4700,  // Acme Corp
  "CUST-1001": 850,   // TechStart Inc
  "CUST-1002": 0      // Global Solutions (no orders yet)
}
```

### 3.2 Orders by Region

To get sales by region, we need to join customer data. Create this formula:

```javascript
// First, get customer regions
=TABLE('Customers') |> PLUCK('region')

// For more complex queries, you can use helper cells
// Put customer lookup in one cell, then reference it
```

Or create a more sophisticated query that groups by product:

```javascript
=TABLE('Orders') |> GROUP_BY('product') |> SUM('amount')
```

Result:
```javascript
{
  "Widget Pro": 1500,
  "Gadget Plus": 850,
  "Super Gizmo": 3200
}
```

### 3.3 High-Value Orders

Filter orders over $2000:

```javascript
=TABLE('Orders') |> WHERE('amount > 2000') |> RESULT()
```

Returns an array of matching rows:
```javascript
[
  ["ORD-5002", "CUST-1000", "2025-01-20", "Super Gizmo", 15, 3200]
]
```

### 3.4 Average Order Value by Product

```javascript
=TABLE('Orders') |> GROUP_BY('product') |> AVG('amount')
```

### 3.5 Count Orders by Customer

```javascript
=TABLE('Orders') |> GROUP_BY('customer_id') |> COUNT()
```

---

## Step 4: Create Customer Lookup Functions

Let's add helper formulas to make it easy to look up customer information.

### 4.1 Customer Name Lookup

In your Analytics sheet, create a lookup function:

```javascript
// In cell A10, enter a customer ID: CUST-1000
// In cell B10, look up the company name:

=INDEX(TABLE('Customers') |> WHERE('customer_id == "' & A10 & '"') |> PLUCK('company'), 1)
```

### 4.2 Customer Contact Info

```javascript
// Get email for a customer
=INDEX(TABLE('Customers') |> WHERE('customer_id == "' & A10 & '"') |> PLUCK('email'), 1)

// Get region for a customer
=INDEX(TABLE('Customers') |> WHERE('customer_id == "' & A10 & '"') |> PLUCK('region'), 1)
```

---

## Step 5: Bulk Data Import

Let's say you receive a CSV export of new orders. You can bulk import them using batch operations.

### 5.1 Prepare Import Data

Suppose you have this data to import:

```javascript
const newOrders = [
  {customer: "CUST-1001", date: "2025-01-25", product: "Widget Pro", qty: 8, amount: 1200},
  {customer: "CUST-1002", date: "2025-01-26", product: "Gadget Plus", qty: 12, amount: 2040},
  {customer: "CUST-1000", date: "2025-01-27", product: "Widget Pro", qty: 20, amount: 3000}
];
```

### 5.2 Convert to Batch Format

```javascript
// Assuming next row is 5 (orders in rows 2-4)
const batchUpdates = [];
let row = 5;

newOrders.forEach(order => {
  batchUpdates.push(
    {address: `B${row}`, value: order.customer},
    {address: `C${row}`, value: order.date},
    {address: `D${row}`, value: order.product},
    {address: `E${row}`, value: String(order.qty)},
    {address: `F${row}`, value: String(order.amount)}
  );
  row++;
});

// Execute batch import
CALL BATCH_SET_CELLS(JSON.stringify(batchUpdates));
```

The order IDs in column A will be automatically generated: `ORD-5003`, `ORD-5004`, `ORD-5005`

---

## Step 6: Create Dashboard Formulas

Now let's create a dashboard with key metrics.

### 6.1 Dashboard Layout

Create a "Dashboard" sheet with this layout:

| A | B |
|---|---|
| **Metric** | **Value** |
| Total Customers | `=COUNT(TABLE('Customers') |> RESULT())` |
| Total Orders | `=COUNT(TABLE('Orders') |> RESULT())` |
| Total Revenue | `=SUM(TABLE('Orders') |> PLUCK('amount'))` |
| Average Order Value | `=AVG(TABLE('Orders') |> PLUCK('amount'))` |
| Top Region by Revenue | (see below) |

### 6.2 Advanced Metrics

**Revenue by Region** (requires joining customer and order data):

Since we need to join data, create a helper column in Orders sheet that looks up the region:

In Orders sheet, column G (Region lookup):
```javascript
=INDEX(TABLE('Customers') |> WHERE('customer_id == "' & B2 & '"') |> PLUCK('region'), 1)
```

Then in Dashboard:
```javascript
// Add region to Orders table definition first
=TABLE('Orders') |> GROUP_BY('region_lookup') |> SUM('amount')
```

**Top Product:**

```javascript
// This is a simplified version; in real use you'd need to sort results
=TABLE('Orders') |> GROUP_BY('product') |> SUM('amount')
```

---

## Step 7: Add Data Entry Forms

Use context-aware validation to create smart data entry forms.

### 7.1 New Customer Form

Create a "New Customer" sheet with a form layout:

| A | B |
|---|---|
| Customer ID | (auto-filled) |
| Company Name | [entry field] |
| Contact Name | [entry field] |
| Email | [entry field with validation] |
| Region | [dropdown] |

Add validation:

```javascript
// Email must be unique (onCreate only)
model.setCellValidation('B4', {
  type: 'contextual',
  onCreate: 'UNIQUE("Customers!D:D", value) && value CONTAINS "@"',
  onUpdate: 'value CONTAINS "@"'  // Less strict on update
});
```

### 7.2 Submit Button Action

When user clicks "Submit", use batch operations to add the customer:

```javascript
// Get form values
const formData = {
  company: model.getCellValue('B2'),
  contact: model.getCellValue('B3'),
  email: model.getCellValue('B4'),
  region: model.getCellValue('B5')
};

// Find next empty row in Customers sheet
const nextRow = findNextEmptyRow('Customers');

// Submit via batch
CALL BATCH_SET_CELLS('[
  {"address": "Customers!B' + nextRow + '", "value": "' + formData.company + '"},
  {"address": "Customers!C' + nextRow + '", "value": "' + formData.contact + '"},
  {"address": "Customers!D' + nextRow + '", "value": "' + formData.email + '"},
  {"address": "Customers!E' + nextRow + '", "value": "' + formData.region + '"}
]')

// Clear form
CALL BATCH_SET_CELLS('[
  {"address": "B2", "value": ""},
  {"address": "B3", "value": ""},
  {"address": "B4", "value": ""},
  {"address": "B5", "value": ""}
]')
```

---

## Step 8: Advanced Queries

Let's explore some advanced querying capabilities.

### 8.1 Multi-Step Queries

Chain multiple WHERE conditions:

```javascript
// High-value West region orders for widgets
=TABLE('Orders')
  .WHERE('amount > 1000')
  .WHERE('product CONTAINS "Widget"')
  .RESULT()
```

### 8.2 Nested Aggregations

```javascript
// Average revenue per customer
// First get total by customer, then average those totals
=AVG(VALUES(TABLE('Orders') |> GROUP_BY('customer_id') |> SUM('amount')))
```

### 8.3 Custom Calculations

```javascript
// Revenue per unit by product
const byProduct = TABLE('Orders') |> GROUP_BY('product');
const revenue = byProduct.SUM('amount');
const quantity = byProduct.SUM('quantity');

// Divide revenue by quantity for each product
=MAPOBJ(revenue, (product, total) => ({
  product: product,
  revenuePerUnit: total / quantity[product]
}))
```

---

## Step 9: Reporting

Create automated reports using query chaining.

### 9.1 Sales Report

Create a "Sales Report" sheet:

**Report Header:**
```
Sales Report - Generated [DATE]
```

**Summary Section:**
```javascript
Total Orders: =COUNT(TABLE('Orders') |> RESULT())
Total Revenue: =SUM(TABLE('Orders') |> PLUCK('amount'))
Average Order: =AVG(TABLE('Orders') |> PLUCK('amount'))
```

**Breakdown by Region:**
```javascript
// Assuming you added region lookup column
=TABLE('Orders') |> GROUP_BY('region_lookup') |> SUM('amount')
```

**Top 5 Products:**
```javascript
=TABLE('Orders') |> GROUP_BY('product') |> SUM('amount')
// Then manually sort or use additional logic
```

### 9.2 Customer Activity Report

```javascript
// Orders per customer
=TABLE('Orders') |> GROUP_BY('customer_id') |> COUNT()

// Revenue per customer
=TABLE('Orders') |> GROUP_BY('customer_id') |> SUM('amount')

// Average order value per customer
=TABLE('Orders') |> GROUP_BY('customer_id') |> AVG('amount')
```

---

## Step 10: Maintenance and Utilities

### 10.1 Find Duplicate Emails

```javascript
// Get all emails
const emails = TABLE('Customers') |> PLUCK('email');

// Find duplicates (manual check or use helper formula)
=FILTER(emails, (email, idx) =>
  FIND(emails, email) != idx
)
```

### 10.2 Update Customer Region

```javascript
// Find customer by email
const customerId = FIND_ROW_BY_ID("CUST-1000");

// Update region
CALL UPDATE_ROW_BY_ID("CUST-1000", '[
  {"column": "E", "value": "South"}
]')
```

### 10.3 Bulk Update Order Status

If you add a "Status" column to orders:

```javascript
// Mark all old orders as "Archived"
CALL BATCH_EXECUTE('[
  {"command": "SETCELL", "args": ["Orders!G2", "Archived"]},
  {"command": "SETCELL", "args": ["Orders!G3", "Archived"]},
  {"command": "SETCELL", "args": ["Orders!G4", "Archived"]}
]')
```

---

## Complete CRM Structure

Here's the final structure of your CRM:

### Sheet 1: Customers
- Columns: CustomerID (auto), CompanyName, ContactName, Email (unique), Region
- Validation: Unique emails, region dropdown
- Table metadata defined

### Sheet 2: Orders
- Columns: OrderID (auto), CustomerID, OrderDate, Product, Quantity, Amount
- Validation: CustomerID must exist, positive numbers
- Table metadata defined

### Sheet 3: Analytics
- Sales by customer
- Sales by product
- Sales by region
- High-value orders
- Average metrics

### Sheet 4: Dashboard
- Key metrics (KPIs)
- Visual summaries
- Trend indicators

### Sheet 5: New Customer Form
- Data entry form
- Context-aware validation
- Batch submission

---

## Key Takeaways

You've learned how to:

1. ✅ **Use Auto-Increment IDs** - Automatically generate unique identifiers
2. ✅ **Define Table Metadata** - Create schemas for better query readability
3. ✅ **Query with TABLE()** - Use column names instead of letters
4. ✅ **Chain Queries** - Build complex filters with WHERE, GROUP_BY, etc.
5. ✅ **Validate Data** - Use context-aware validation (onCreate vs onUpdate)
6. ✅ **Batch Operations** - Efficiently import and update bulk data
7. ✅ **Build Analytics** - Create powerful reports with aggregations
8. ✅ **Join Data** - Link related tables (customers and orders)

---

## Next Steps

- **Add Products Table**: Create a products catalog with pricing
- **Implement Inventory**: Track product stock levels
- **Create Invoicing**: Generate invoices from orders
- **Add User Management**: Track who created/modified records
- **Build Charts**: Visualize sales trends over time
- **Export Reports**: Use Control Bus to generate PDF reports
- **API Integration**: Connect to external CRM systems via Control Bus

---

## Additional Resources

- **[FEATURES.md](FEATURES.md)** - Complete feature documentation
- **[README.md](../README.md)** - Getting started guide
- **[TESTING-CONTROL-BUS.md](../TESTING-CONTROL-BUS.md)** - Control Bus API reference
- **[ENHANCEMENT_PLAN.md](../ENHANCEMENT_PLAN.md)** - Technical implementation details

---

**Congratulations!** You've built a functional CRM using RexxJS Spreadsheet's advanced features. This demonstrates the power of combining spreadsheet simplicity with database-like capabilities.
