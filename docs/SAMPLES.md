# Sample databases

Rutba Data Files comes with a working database for each kind of small business we make Office for.
Each one has related tables, the questions that business asks as saved queries, a form for every table and
reports to print, and it comes either with a year or more of invented records to learn from, or empty, ready
for your own. On the Data Files start screen, choose a business and then **With sample records** or **Empty**.
The same files are on [office.rutba.io/samples](https://office.rutba.io/samples) to download.

Everything in the records is made up: the businesses, the people, their addresses, e-mail addresses on
example domains and phone numbers from the ranges kept for drama and fiction. The dates count back from the
day the file is made, so "due this week" always has something in it. This page was made on 2026-10-21.

The samples are made by `packages/database/src/samples` (one file each), checked by `check.js` and
`tests/database-samples.test.js`, and written as files by `node tools/make-samples.mjs`. This page is made by
`node tools/make-samples.mjs --docs` from each sample's own description, so it says what the files hold.

With each release, the downloads on office.rutba.io are made again so their dates are current: in the
management repository, `npm run import:samples` in `portal/apps/office` builds every sample from this
repository, zips it with a read-me into `public/samples/` and writes the pages' catalogue,
`src/site/samples.json`. A sample that fails `checkSample` stops the import.

| Sample | For | Tables | Queries | Reports |
|---|---|---|---|---|
| [Northbank Trading](#northbank-trading) | Wholesale and trading | 11 | 11 | 5 |
| [Larkspur Corner Store](#larkspur-corner-store) | Shop | 10 | 12 | 7 |
| [Cedar Lane Clinic](#cedar-lane-clinic) | Clinic | 9 | 17 | 6 |
| [Marigold Hair and Beauty](#marigold-hair-and-beauty) | Salon | 10 | 15 | 7 |
| [Ironbridge Repairs](#ironbridge-repairs) | Repair shop | 10 | 18 | 6 |
| [Bramble Studio](#bramble-studio) | Consultancy or agency | 10 | 17 | 7 |
| [Little Fig Cafe](#little-fig-cafe) | Cafe or restaurant | 13 | 20 | 9 |
| [Hollybrook Lettings](#hollybrook-lettings) | Property rentals | 9 | 18 | 8 |
| [Lantern Learning Centre](#lantern-learning-centre) | School or training centre | 13 | 13 | 8 |
| [Rowanmere Community Trust](#rowanmere-community-trust) | Charity or club | 13 | 20 | 8 |
| [Granite and Grain Woodworks](#granite-and-grain-woodworks) | Workshop or maker | 14 | 17 | 9 |

## Northbank Trading

**Wholesale and trading.** Customers, suppliers, products, orders, purchase orders, stock and invoices for a wholesaler.

Northbank Trading is an invented wholesaler of specialty foods and supplies for the trade: teas and coffee, spices, rice and grains, dried fruit and nuts, oils, sauces, baking goods and kitchen supplies. It sells to shops, delis, cafes, restaurants and hotels in ten countries, and buys from a dozen suppliers.

The sample holds eighteen months of trading, made up a day at a time. Customers order more as the year ends. Each order is cut to what is on the shelf, so stock never goes below nought, and it goes down as orders ship. Twice a week the buyer raises purchase orders for whatever has fallen to its reorder level, and those arrive some days later. Every shipped order is invoiced; most invoices are paid on time, some late, and the newest are not yet due.

Customers, suppliers and carriers share one Companies table, told apart by their Kind, so a company that both buys from you and sells to you is entered once.

It suits, with small changes: A distributor or importer; A cash and carry; A builders' merchant or parts supplier; A wine merchant selling to the trade; An online shop that sells to businesses; A manufacturer's sales office.

### Tables

#### Companies

Every business you deal with: customers, suppliers and the carriers who deliver for you. The sample has 63 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Name | Short text, required | The company's name. |
| Kind | Short text | Customer, Supplier or Shipper. |
| Address | Short text | The street address. |
| City | Short text | The town or city. |
| Postcode | Short text | The postcode or ZIP code. |
| Country | Short text | The country. |
| Phone | Short text | The main phone number. |
| E-mail | Short text | Where orders or invoices go. |
| Payment Days | Whole number | How many days after the invoice payment is due: 30 means thirty days. |
| Credit Limit | Currency | The most a customer may owe at once. |
| Notes | Long text | Anything worth remembering about the company. |

#### Contacts

The people at each company. The sample has 84 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Company | Points at Companies, required | The company they work for. |
| First Name | Short text | Their first name. |
| Last Name | Short text, required | Their family name. |
| Job Title | Short text | Their job, such as Buyer or Account Manager. |
| E-mail | Short text | Their e-mail address. |
| Phone | Short text | Their phone number. |

#### Employees

Your own staff: sales, buying, the warehouse and accounts. The sample has 9 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| First Name | Short text | Their first name. |
| Last Name | Short text, required | Their family name. |
| Job Title | Short text | Their job. |
| E-mail | Short text | Their work e-mail address. |
| Phone | Short text | Their work phone number. |
| Hired | Date | The day they joined. |

#### Categories

The groups your products are sold in. The sample has 8 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Name | Short text, required, unique | The group's name. Each name is used once. |
| Description | Short text | What the group holds. |

#### Products

What you sell, one record for each line in the catalogue. The sample has 55 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Code | Short text, unique | Your own product code. Each code is used once. |
| Name | Short text, required | The product as it appears on the catalogue and invoices. |
| Category | Points at Categories | The group it is sold in. |
| Supplier | Points at Companies | Who you buy it from. |
| Unit | Short text | How it is sold, such as a 10 kg sack or a box of 100. |
| Cost Price | Currency | What one unit costs you. |
| List Price | Currency | What one unit sells for before any discount. |
| Units in Stock | Whole number | How many units are in the warehouse now. |
| Reorder Level | Whole number | When stock falls to this level, the product is ordered again. |
| Reorder Quantity | Whole number | How many to order each time. |
| Discontinued | Yes/No | Ticked for a product you no longer sell; old orders keep it. |

#### Orders

Customers' orders, one record each. The sample has 635 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Customer | Points at Companies, required | Who ordered. |
| Employee | Points at Employees | The sales person who took it. |
| Order Date | Date | The day it was placed. |
| Required Date | Date | When the customer needs it. |
| Shipped Date | Date | The day it left the warehouse; empty until then. |
| Shipper | Points at Companies | The carrier taking it. |
| Shipping Fee | Currency | What the customer pays for delivery. |
| Status | Short text | New (not shipped yet), Shipped, Closed (shipped and paid) or Cancelled. |
| Notes | Long text | Delivery instructions or anything else. |

#### Order Lines

The products on each order. The sample has 1,883 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Order | Points at Orders, required | The order the line belongs to. |
| Product | Points at Products, required | What was ordered. |
| Quantity | Whole number | How many units. |
| Unit Price | Currency | The price of one unit on this order, kept so a later price change does not alter it. |
| Discount | Number | The discount as a fraction: 0.05 is five percent. |

#### Purchase Orders

Orders you place with your suppliers. The sample has 197 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Supplier | Points at Companies, required | Who you ordered from. |
| Employee | Points at Employees | Who placed the order. |
| Ordered On | Date | The day it was sent. |
| Expected | Date | When the goods should arrive. |
| Received On | Date | The day they arrived; empty until then. |
| Status | Short text | Submitted or Received. |

#### Purchase Lines

The products on each purchase order. The sample has 270 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Purchase Order | Points at Purchase Orders, required | The purchase order the line belongs to. |
| Product | Points at Products, required | What was ordered. |
| Quantity | Whole number | How many units. |
| Unit Cost | Currency | What one unit costs on this order. |

#### Stock Movements

Every change to stock: goods received, goods sold and corrections from a stock count. Added up, they give what is on the shelf. The sample has 2,077 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Product | Points at Products, required | The product that moved. |
| When | Date | The day it moved. |
| Kind | Short text | Received, Sold or Adjusted. |
| Quantity | Whole number | How many: above nought coming in, below nought going out. |
| Reference | Short text | The order or purchase order it came from, or Stock count. |

#### Invoices

An invoice for each shipped order. The sample has 608 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Order | Points at Orders, required | The order invoiced. |
| Invoice Date | Date | The day it was raised, the day the order shipped. |
| Due Date | Date | When payment is due, from the customer's payment days. |
| Amount | Currency | What the invoice comes to, delivery included. |
| Paid On | Date | The day it was paid; empty while it is owed. |

### Relationships

- Contacts.Company points at Companies.ID
- Products.Category points at Categories.ID
- Products.Supplier points at Companies.ID
- Orders.Customer points at Companies.ID
- Orders.Employee points at Employees.ID
- Orders.Shipper points at Companies.ID
- Order Lines.Order points at Orders.ID
- Order Lines.Product points at Products.ID
- Purchase Orders.Supplier points at Companies.ID
- Purchase Orders.Employee points at Employees.ID
- Purchase Lines.Purchase Order points at Purchase Orders.ID
- Purchase Lines.Product points at Products.ID
- Stock Movements.Product points at Products.ID
- Invoices.Order points at Orders.ID

### Queries

#### Order totals

Every order with its customer, status and the value of its goods after discounts, the newest first.

From Orders, Companies, Order Lines. Shows Order, Order Date, Customer, Status, Goods.
Calculated: Goods is `Round([Quantity] * [Unit Price] * (1 - Nz([Discount])), 2)`.

#### Sales by month

Sales and units for each month, cancelled orders left out.

From Orders, Order Lines. Shows Month, Sales, Units.
Calculated: Month is `Format([Order Date], "yyyy-mm")`; Sales is `Round([Quantity] * [Unit Price] * (1 - Nz([Discount])), 2)`.
Only where `Status <> "Cancelled"`.

#### Sales by category

What each product group has sold, the biggest first.

From Order Lines, Orders, Products, Categories. Shows Category, Sales, Units.
Calculated: Sales is `Round([Quantity] * [Unit Price] * (1 - Nz([Discount])), 2)`.
Only where `Status <> "Cancelled"`.

#### Sales by employee

What each sales person has sold.

From Orders, Employees, Order Lines. Shows Employee, Sales.
Calculated: Employee is `[First Name] & " " & [Last Name]`; Sales is `Round([Quantity] * [Unit Price] * (1 - Nz([Discount])), 2)`.
Only where `Status <> "Cancelled"`.

#### Top customers

Customers by what they have bought, with their country.

From Companies, Orders, Order Lines. Shows Customer, Country, Sales.
Calculated: Sales is `Round([Quantity] * [Unit Price] * (1 - Nz([Discount])), 2)`.
Only where `Status <> "Cancelled"`.

#### Products to reorder

Products at or below their reorder level, by supplier, with how many to order. Discontinued lines are left out.

From Products, Companies. Shows Supplier, Code, Product, Units in Stock, Reorder Level, Reorder Quantity.
Only where `Above Reorder Level <= 0` and `Discontinued = No`.

#### Orders to ship

Orders not shipped yet, oldest first, with the days left before the customer needs them.

From Orders, Companies. Shows Order, Order Date, Customer, Required Date, Days Left.
Calculated: Days Left is `DateDiff("d", Date(), [Required Date])`.
Only where `Shipped Date Is Null` and `Status <> "Cancelled"`.

#### Unpaid invoices

Invoices not yet paid, by due date, with how many days each is overdue (nought when it is not due yet).

From Invoices, Orders, Companies. Shows Invoice, Customer, Invoice Date, Due Date, Amount, Days Overdue.
Calculated: Days Overdue is `IIf(DateDiff("d", [Due Date], Date()) > 0, DateDiff("d", [Due Date], Date()), 0)`.
Only where `Paid On Is Null`.

#### Stock value

Units in stock and their value at cost, by product group.

From Products, Categories. Shows Category, Units, Value at Cost.
Calculated: Value at Cost is `Round([Units in Stock] * [Cost Price], 2)`.

#### Product list

The products you still sell, by group and code, with their list prices.

From Products, Categories. Shows Category, Code, Product, Unit, List Price.
Only where `Discontinued = No`.

#### Purchases on order

Purchase orders sent and not yet received, by the day they are expected.

From Purchase Orders, Companies. Shows Purchase Order, Supplier, Ordered On, Expected, Status.
Only where `Status = Submitted`.

### Reports

#### Product catalogue

The price list: every product you still sell by group, with its unit and list price.

From the query Product list, grouped by Category: Code, Product, Unit, List Price.

#### Monthly sales

One line a month with units and sales, and the totals.

From the query Sales by month: Month, Units, Sales, with total of sales, total of units.

#### Orders by customer

Each customer's orders, the newest first, with a total for each customer.

From the query Order totals, grouped by Customer: Order, Order Date, Status, Goods, with total of goods.

#### Invoices due

Unpaid invoices by customer with their due dates, days overdue and a total owed: the list for chasing payment.

From the query Unpaid invoices, grouped by Customer: Invoice, Invoice Date, Due Date, Days Overdue, Amount, with total of amount.

#### Stock to reorder

What to order, on a page for each supplier.

From the query Products to reorder, grouped by Supplier: Code, Product, Units in Stock, Reorder Level, Reorder Quantity.

### Starting empty

Start with Categories and Companies (your suppliers and carriers first, then your customers), then Employees and Products. Orders, purchase orders, stock movements and invoices come in as you trade. If you already hold stock, enter it as one Received movement per product and set Units in Stock to match.

### Making it yours

- Sell in more than one currency: add a Currency field to Companies and a rate to Orders, and a calculated column for the amount in your own currency.
- Keep several warehouses: add a Warehouses table, point Stock Movements at it, and group the stock queries by warehouse.
- Record part payments: add a Payments table pointing at Invoices, and an Unpaid invoices query that sums what has been paid against each one.
- Give customers their own prices: add a Price Lists table (customer, product, price) and use it when taking an order.
- Track batches and best-before dates: add a Batch and a Best Before field to Purchase Lines and Stock Movements.
- Add sales tax: add a Tax Rate field to Products and a calculated column for the tax on each line.

## Larkspur Corner Store

**Shop.** Products, the till, stock, deliveries, loyalty customers and running costs for a neighbourhood shop.

Larkspur Corner Store is an invented neighbourhood shop open from seven in the morning to nine at night, every day of the week. It sells bread, milk, fruit and vegetables, drinks, snacks, groceries, household goods and newspapers, from about forty lines bought from six suppliers.

The sample holds a year of trading made up a day at a time. The till rings up a sale every few minutes, more at lunchtime and after school and more on Saturdays. Each supplier delivers on its own days for what has run low, so stock goes down between deliveries and back up again. At each month end the rent, wages, power and card fees are paid, so the takings can be set against what the shop costs to run.

Use it to see how a till, stock and suppliers fit together in one database, or start from it empty for a shop of your own.

It suits, with small changes: A convenience store or minimarket; A farm shop; A newsagent; A gift shop; A pharmacy counter for the goods it sells off the shelf; A market stall.

### Tables

#### Categories

The sections of the shop, used to group products on reports. The sample has 9 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Name | Short text, required, unique | The section, such as Bakery or Drinks. Each name is used once. |

#### Suppliers

The wholesalers and producers the shop buys from. The sample has 6 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Name | Short text, required | The supplier's name. |
| Contact | Short text | The person you deal with there. |
| Phone | Short text | Their phone number. |
| E-mail | Short text | Where orders and queries go. |
| Delivers On | Short text | The days the supplier comes, so you know when an order will arrive. |

#### Products

Everything on the shelves, one record per line you sell. The sample has 41 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Barcode | Short text, unique | The number on the pack the till scans. Each barcode is used once. |
| Name | Short text, required | The product as it shows on the till and the receipt. |
| Category | Points at Categories | The section it belongs to. |
| Supplier | Points at Suppliers | Who you buy it from. |
| Cost Price | Currency | What one costs you to buy. |
| Sale Price | Currency | What one sells for. |
| In Stock | Whole number | How many are on the shelf and in the back now. |
| Reorder Level | Whole number | When stock falls to this number or below, it goes on the shopping list. |
| Active | Yes/No | Untick a product you no longer sell; it stays in the old sales. |

#### Staff

The people who work the till. The sample has 4 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| First Name | Short text | Their first name. |
| Last Name | Short text, required | Their family name. |
| Role | Short text | Owner, shop assistant or weekend assistant. |
| Phone | Short text | A number to reach them. |
| Started | Date | The day they started. |

#### Customers

Regulars who signed up for the loyalty card. The sample has 60 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| First Name | Short text | Their first name. |
| Last Name | Short text, required | Their family name. |
| Phone | Short text | Their mobile number. |
| E-mail | Short text | For offers, if they asked for them. |
| Joined | Date | The day they signed up. |
| Points | Whole number | Loyalty points earned: one for each whole amount spent. |

#### Sales

Each time the till is rung up: one record per receipt. The sample has 21,535 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| When | Date and time | The date and time of the sale. |
| Staff | Points at Staff | Who was on the till. |
| Customer | Points at Customers | The loyalty customer, when they showed their card. |
| Payment | Short text | Card, Cash or Mobile. |
| Refunded | Yes/No | Ticked when the whole sale was refunded; refunded sales are left out of the takings. |

#### Sale Lines

The items on each receipt. The sample has 56,648 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Sale | Points at Sales, required | The receipt the line belongs to. |
| Product | Points at Products, required | What was sold. |
| Quantity | Whole number | How many. |
| Unit Price | Currency | The price of one at the time, kept so a later price change does not alter old sales. |
| Discount | Number | Any discount, as a fraction: 0.1 is ten percent off. |

#### Deliveries

Each delivery from a supplier, with its invoice. The sample has 854 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Supplier | Points at Suppliers, required | Who delivered. |
| Received | Date | The day it came. |
| Invoice Number | Short text | The number on the supplier's invoice. |
| Paid | Yes/No | Ticked once the invoice is paid. |

#### Delivery Lines

What came in each delivery. The sample has 4,154 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Delivery | Points at Deliveries, required | The delivery the line belongs to. |
| Product | Points at Products, required | What came. |
| Quantity | Whole number | How many. |
| Unit Cost | Currency | What one cost on this invoice. |

#### Expenses

What the shop costs to run, other than the goods it sells. The sample has 70 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Date | Date | The day it was paid. |
| Kind | Short text | Rent, Wages, Utilities, Card fees, Insurance, Phone and internet or Repairs. |
| Description | Short text | What it was for. |
| Amount | Currency | How much. |

### Relationships

- Products.Category points at Categories.ID
- Products.Supplier points at Suppliers.ID
- Sales.Staff points at Staff.ID
- Sales.Customer points at Customers.ID
- Sale Lines.Sale points at Sales.ID
- Sale Lines.Product points at Products.ID
- Deliveries.Supplier points at Suppliers.ID
- Delivery Lines.Delivery points at Deliveries.ID
- Delivery Lines.Product points at Products.ID

### Queries

#### Today's takings

What the till has taken today so far, by how people paid, and how many items went out. Run it at closing time to check against the cash drawer and the card terminal.

From Sales, Sale Lines. Shows Payment, Takings, Items.
Calculated: Takings is `Round([Quantity] * [Unit Price] * (1 - Nz([Discount])), 2)`.
Only where `Days Ago = 0` and `Refunded = No`.

#### Daily takings

Takings and items sold on each of the last 30 days, the newest first, to spot a slow day.

From Sales, Sale Lines. Shows Day, Takings, Items.
Calculated: Day is `Format([When], "yyyy-mm-dd")`; Takings is `Round([Quantity] * [Unit Price] * (1 - Nz([Discount])), 2)`.
Only where `Days Ago Between 0 And 30` and `Refunded = No`.

#### Monthly takings

Takings for each month of the year, refunds left out.

From Sales, Sale Lines. Shows Month, Takings.
Calculated: Month is `Format([When], "yyyy-mm")`; Takings is `Round([Quantity] * [Unit Price] * (1 - Nz([Discount])), 2)`.
Only where `Refunded = No`.

#### Busy hours

How many sales the till rang up in each hour of the day over the last 90 days, to plan who works when.

From Sales. Shows Hour, Sales.
Calculated: Hour is `Mid([When], 12, 2)`.
Only where `Days Ago Between 0 And 90`.

#### Best sellers

The products that sold most in the last 30 days, with what they took.

From Sale Lines, Products, Sales. Shows Product, Sold, Takings.
Calculated: Takings is `Round([Quantity] * [Unit Price] * (1 - Nz([Discount])), 2)`.
Only where `Refunded = No` and `Days Ago Between 0 And 30`.

#### Profit by category

For each section of the shop, what it took and what it earned after the cost of the goods, the best earner first.

From Sale Lines, Products, Categories, Sales. Shows Category, Takings, Profit.
Calculated: Takings is `Round([Quantity] * [Unit Price] * (1 - Nz([Discount])), 2)`; Profit is `Round([Quantity] * ([Unit Price] * (1 - Nz([Discount])) - [Cost Price]), 2)`.
Only where `Refunded = No`.

#### Takings by staff

What each person rang up over the last 30 days.

From Sales, Staff, Sale Lines. Shows Staff Member, Role, Takings.
Calculated: Staff Member is `[First Name] & " " & [Last Name]`; Takings is `Round([Quantity] * [Unit Price] * (1 - Nz([Discount])), 2)`.
Only where `Days Ago Between 0 And 30` and `Refunded = No`.

#### Low stock

Products at or below their reorder level, by supplier, with the days each supplier delivers. This is the shopping list.

From Products, Suppliers. Shows Supplier, Delivers On, Product, In Stock, Reorder Level.
Only where `Spare <= 0` and `Active = Yes`.

#### Stock value

What the stock on the shelves is worth by section, at cost and at sale price.

From Products, Categories. Shows Category, Items, At Cost, At Sale Price.
Calculated: At Cost is `Round([In Stock] * [Cost Price], 2)`; At Sale Price is `Round([In Stock] * [Sale Price], 2)`.

#### Costs by month

What the shop paid out each month, by kind of cost.

From Expenses. Shows Month, Kind, Amount.
Calculated: Month is `Format([Date], "yyyy-mm")`.

#### Best customers

Loyalty customers by what they have spent, with their points.

From Customers, Sales, Sale Lines. Shows Customer, Spent, Points.
Calculated: Customer is `[First Name] & " " & [Last Name]`; Spent is `Round([Quantity] * [Unit Price] * (1 - Nz([Discount])), 2)`.
Only where `Refunded = No`.

#### Unpaid deliveries

Supplier invoices not yet paid, oldest first, with what each one comes to.

From Deliveries, Suppliers, Delivery Lines. Shows Delivery, Supplier, Received, Invoice Number, Owed.
Calculated: Owed is `Round([Quantity] * [Unit Cost], 2)`.
Only where `Paid = No`.

### Reports

#### Takings by month

One line a month with the takings and a total for the year.

From the query Monthly takings: Month, Takings, with total of takings.

#### Takings, the last 30 days

Takings and items for each of the last 30 days, with the total and the average day.

From the query Daily takings: Day, Items, Takings, with total of takings, average of takings.

#### Profit by category

Takings and profit by section, with totals: the page to look at when deciding what to stock more of.

From the query Profit by category: Category, Takings, Profit, with total of takings, total of profit.

#### To order

What to order, on a page for each supplier, with what is left and the reorder level.

From the query Low stock, grouped by Supplier: Product, In Stock, Reorder Level.

#### Stock take sheet

Every product by section with its barcode and what the database says is in stock, to count against.

From the table Products, grouped by Category: Name, Barcode, In Stock, with total of in stock.

#### Running costs

Each month's costs by kind, with a total for the month.

From the query Costs by month, grouped by Month: Kind, Amount, with total of amount.

#### Deliveries to pay

Supplier invoices still to pay, by supplier, with a total for each.

From the query Unpaid deliveries, grouped by Supplier: Delivery, Received, Invoice Number, Owed, with total of owed.

### Starting empty

Start with Categories and Suppliers, then add your Products with their prices and reorder levels, then your Staff. Sales and Deliveries come in as you trade. Put your running costs in Expenses as you pay them.

### Making it yours

- Sell by weight: add a Unit field to Products (each, kg, 100 g) and make Quantity in Sale Lines a Number rather than a whole number.
- Track sell-by dates: add a Best Before field to Delivery Lines and a query for lines within three days of it.
- Add VAT or sales tax: add a Tax Rate field to Categories and a calculated column in the takings queries for the tax in each line.
- Keep more than one till or shop: add a Tills table, point Sales at it, and group the takings queries by till.
- Reward points differently: change Points to count one for every two spent, or add a Points Redeemed field to Sales.
- Order from the shopping list: copy the Low stock query into Worksheets and send it to the supplier, or add an Orders table for what you asked for and when.

## Cedar Lane Clinic

**Clinic.** Patients, appointments, visit notes, invoices and recalls for a small private clinic.

For a small clinic that books appointments, writes a short note at each visit and bills patients or their insurers. Cedar Lane Clinic is invented: two doctors, a practice nurse and two physiotherapists, with a reception desk that takes bookings by phone and online.

The records cover the year up to the day the file is made, and the weeks after it. Patients book a few days ahead. Some cancel and a few do not turn up, and the clinic charges a fee for some of those. A physiotherapy assessment leads to a course of weekly sessions, a health check sets a recall for a year later, and the doctors are busier in winter while the nurse is busiest in the autumn when flu vaccinations are given.

Each visit has an invoice, made out to the patient or, when the policy covers the service, to the insurer. Most patients pay at the desk; insurers pay within their terms, one of them slowly, so there is always something to chase. The saved queries answer what a practice manager asks each day and each month: who is booked this week, how full each diary is, who missed an appointment, who owes what, what came in, and which patients are due to be recalled.

Every patient and every note in it is made up. The notes are short and ordinary on purpose; a real clinic keeps its own wording and its own rules on who may read them.

It suits, with small changes: A family doctor or GP practice that sees private patients; A physiotherapy or sports injury practice; A chiropractic or osteopathy clinic; A dental surgery, with treatments in place of services; A podiatry or foot care clinic; A veterinary practice, with pets as patients and owners for contact details; A counselling or talking therapy practice.

### Tables

#### Insurers

The health insurers whose policies the clinic accepts, and one record called Self-pay for patients who pay for themselves. Every patient points at one of these, so a patient list can always say who pays. The sample has 5 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Name | Short text, required, unique | The insurer as it appears on claims and on the patient list. The Self-pay record stands for no insurer. |
| Kind | Short text | Insurer for a real insurer, Self-pay for the record that stands for paying yourself. Queries use it to tell the two apart. |
| Phone | Short text | The claims line the reception desk rings about a claim. |
| E-mail | Short text | Where claims and questions about them are sent. |
| Payment Days | Whole number | How many days the insurer takes to pay a claim under its agreement with the clinic. An invoice to the insurer falls due this many days after the visit. |
| Notes | Long text | Anything reception should know: what a policy usually covers, or that an insurer wants a referral letter first. |

#### Patients

One record per patient registered with the clinic, with contact details, who pays for their care and whether they agree to reminders. The sample has 841 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Patient No | Short text, unique | The clinic's own number for the patient, printed on day sheets and letters. It is not a national health number. |
| First Name | Short text | The patient's first name. |
| Last Name | Short text, required | The patient's family name. Lists are sorted by it. |
| Date of Birth | Date | Used to tell two patients with the same name apart and to see who is due a check for their age. |
| Address | Short text | The street address, for recall letters and invoices. |
| City | Short text | The town or city. |
| Postcode | Short text | The postcode, for letters. |
| Phone | Short text | The number reception rings to confirm or move an appointment. |
| E-mail | Short text | Where appointment reminders and invoices are sent, if the patient has given one. |
| Insurer | Points at Insurers, required | Who pays: the patient's insurer, or Self-pay. |
| Policy Reference | Short text | The patient's policy or membership reference with the insurer, quoted on every claim. Empty for self-pay patients. |
| Registered | Date | The day the patient first registered. The new patients query counts these by month. |
| Reminders | Yes/No | Ticked when the patient agrees to reminders and recall letters. Patients without it are rung instead, or not contacted. |
| Notes | Long text | Practical things reception should know before a visit, such as needing step-free access. Not clinical notes, which go in Visit Notes. |

#### Practitioners

The people who see patients: the doctors, the nurse and the physiotherapists. Reception staff are not here because they do not hold a diary. The sample has 5 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| First Name | Short text | The practitioner's first name. |
| Last Name | Short text, required | The practitioner's family name. |
| Known As | Short text, required | The name used on day sheets and letters, such as Dr and the name, or the name and Physio. |
| Role | Short text | Doctor, Nurse or Physiotherapist. Services say which role gives them. |
| E-mail | Short text | The practitioner's clinic e-mail address. |
| Phone | Short text | The practitioner's extension or direct line. |
| Works | Short text | The days and hours the practitioner sees patients, in words. |
| Hours per Week | Number | The hours a week the practitioner has for appointments. The diary query measures bookings against it. |
| Started | Date | The day the practitioner started at the clinic. |

#### Services

What the clinic offers, how long each takes in the diary and what it costs. The sample has 14 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Name | Short text, required, unique | The service as it appears in the diary and on invoices. |
| Kind | Short text | A group of services: Consultation, Follow-up, Health check, Nursing, Vaccination or Physiotherapy. |
| Given By | Short text | The role that gives it: Doctor, Nurse or Physiotherapist. |
| Minutes | Whole number | How long the appointment lasts. The diary is kept in quarter hours, so these are multiples of 15. |
| Price | Currency | The fee for one appointment. |
| Insurer Pays | Yes/No | Ticked when insurers usually cover the service. Health checks and vaccinations are paid by the patient even when insured. |
| Recall Months | Whole number | How many months after this service the patient should come back for it, such as 12 for a yearly health check. Empty when there is no recall. |

#### Appointments

Every booking in the diary, past and future, with what came of it. The sample has 5,343 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Patient | Points at Patients, required | Who the appointment is for. |
| Practitioner | Points at Practitioners, required | Whose diary it is in. |
| Service | Points at Services, required | What the appointment is for, which sets its length and price. |
| Start | Date and time, required | The date and time the appointment starts. |
| Status | Short text | Booked until the day, then Attended or No-show. Cancelled when the patient or the clinic called it off. |
| Booked On | Date | The day the booking was made. |
| Booked Via | Short text | How it was booked: Phone, Online, At the desk, or Recall when it came from a recall letter or call. |
| Cancelled On | Date | The day the appointment was cancelled. A cancellation the day before or on the day counts as late. |

#### Visit Notes

A short note written by the practitioner after each attended appointment. In this sample they are invented and kept to everyday words. The sample has 4,632 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Appointment | Points at Appointments, required | The visit the note is about, which gives the patient, the practitioner and the date. |
| Written | Date and time | When the note was written, usually just after the appointment ended. |
| Reason | Short text | Why the patient came, in a few words. |
| Note | Long text | What was found and what was done or advised. |
| Plan | Short text | What happens next: nothing, a follow-up, a test, a course of sessions or a recall. |

#### Invoices

What the clinic charges, one invoice per visit, and for a missed appointment or a late cancellation when a fee is charged. The sample has 4,846 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Appointment | Points at Appointments, required | The appointment the invoice is for. |
| Invoice Date | Date | The day the invoice was raised: the day of the visit, or of the missed or cancelled appointment. |
| Due Date | Date | When payment is due: 14 days for patients, the insurer's payment days for claims. |
| Bill To | Short text | Patient or Insurer. An insurer is billed when the patient is insured and the service is one insurers pay for. |
| For | Short text | Visit, Missed appointment or Late cancellation. |
| Amount | Currency | The amount charged. |
| Amount Paid | Currency | How much has been paid so far, the sum of the payments against it. |
| Paid On | Date | The day the invoice was paid in full. Empty while anything is still owed. |

#### Payments

Money received against invoices. An invoice can have more than one payment when a patient pays in parts. The sample has 4,603 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Invoice | Points at Invoices, required | The invoice the payment is for. |
| Paid On | Date | The day the money came in. |
| Amount | Currency | How much was paid. |
| Method | Short text | Card, Cash or Bank transfer from a patient, or Insurer for a claim paid by an insurer. |
| Reference | Short text | The insurer's remittance reference or the reference on a bank transfer. |

#### Recalls

Patients the clinic should ask back: for a yearly health check, a blood pressure review, the autumn flu vaccination or a review after physiotherapy. The sample has 746 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Patient | Points at Patients, required | Who to recall. |
| Service | Points at Services, required | The service they are being asked back for. |
| Due Date | Date | When the recall falls due. |
| Reason | Short text | Why, in words for the letter or the call. |
| Sent On | Date | The day a recall letter or reminder went out. Letters go each Monday for recalls due within two weeks. |
| Booked On | Date | The day the patient booked in response. Empty while the recall is still open. |

### Relationships

- Patients.Insurer points at Insurers.ID
- Appointments.Patient points at Patients.ID
- Appointments.Practitioner points at Practitioners.ID
- Appointments.Service points at Services.ID
- Visit Notes.Appointment points at Appointments.ID
- Invoices.Appointment points at Appointments.ID
- Payments.Invoice points at Invoices.ID
- Recalls.Patient points at Patients.ID
- Recalls.Service points at Services.ID

### Queries

#### Next 7 days

For each day of the coming week and each practitioner, how many appointments are booked and the hours they fill, with the first and last start times. Reception uses it to see the shape of the week at a glance.

From Appointments, Practitioners, Services. Shows Day, Practitioner, Appointments, Booked Hours, First Start, Last Start.
Calculated: Day is `Format([Start], "yyyy-mm-dd")`; Booked Hours is `[Minutes] / 60`; First Start is `Mid([Start], 12, 5)`; Last Start is `Mid([Start], 12, 5)`.
Only where `Status = Booked` and `Days Ahead Between 0 And 6`.

#### Day sheet

Every booked appointment for the next seven days, by day and practitioner, with the patient's number, phone, who pays and any practical notes. It is what each practitioner and the desk work from each morning.

From Appointments, Patients, Practitioners, Services, Insurers. Shows Sheet, Time, Patient, Patient No, Phone, Service, Minutes, Paid By, Patient Notes.
Calculated: Sheet is `Format([Start], "yyyy-mm-dd") & " " & [Known As]`; Time is `Mid([Start], 12, 5)`; Patient is `[Patients].[First Name] & " " & [Patients].[Last Name]`.
Only where `Status = Booked` and `Days Ahead Between 0 And 6`.

#### Diary next 14 days

How full each practitioner's diary is for the next two weeks: appointments, hours booked, hours open and the percentage full. Use it to see who can take more patients or where to offer a cancellation slot.

From Appointments, Practitioners, Services. Shows Practitioner, Role, Appointments, Booked Hours, Open Hours, Percent Full.
Calculated: Booked Hours is `[Minutes] / 60`; Open Hours is `[Hours per Week] * 2`; Percent Full is `Round([Minutes] * 100 / ([Hours per Week] * 120), 2)`.
Only where `Status = Booked` and `Days Ahead Between 0 And 13`.

#### No-shows and late cancellations

Appointments in the last 90 days the patient did not attend, or cancelled the day before or on the day, with the phone number. Use it to decide on fees and to talk to patients who miss often.

From Appointments, Patients, Practitioners, Services. Shows Start, Patient, Phone, Practitioner, Service, Status, Cancelled On, Days Notice.
Calculated: Patient is `[Patients].[First Name] & " " & [Patients].[Last Name]`; Days Notice is `DateDiff("d", [Cancelled On], Format([Start], "yyyy-mm-dd"))`.
Only where `Status = No-show` and `Status = Cancelled` and `Days Notice <= 1` and `Days Ago Between 0 And 90` and `Days Ago Between 0 And 90`.

#### Unpaid invoices

Every invoice not yet paid in full, with who owes it, what has been paid and how many days overdue it is. A minus figure means it is not due yet.

From Invoices, Appointments, Patients, Insurers. Shows Invoice, Owed By, Bill To, Patient, For, Invoice Date, Due Date, Amount, Paid So Far, Balance, Days Overdue.
Calculated: Owed By is `IIf([Bill To] = "Insurer", [Insurers].[Name], [Patients].[First Name] & " " & [Patients].[Last Name])`; Patient is `[Patients].[First Name] & " " & [Patients].[Last Name]`; Balance is `Round([Amount] - Nz([Amount Paid]), 2)`; Days Overdue is `DateDiff("d", [Due Date], Date())`.
Only where `Paid On Is Null`.

#### Money owed

The unpaid balance added up for each insurer and each patient, with the number of invoices and the oldest due date. Start here when chasing payments.

From Invoices, Appointments, Patients, Insurers. Shows Owed By, Bill To, Invoices, Balance, Oldest Due.
Calculated: Owed By is `IIf([Bill To] = "Insurer", [Insurers].[Name], [Patients].[First Name] & " " & [Patients].[Last Name])`; Balance is `Round([Amount] - Nz([Amount Paid]), 2)`.
Only where `Paid On Is Null`.

#### Takings by month

Money received each month, split into what came from patients and what came from insurers.

From Payments. Shows Month, From Patients, From Insurers, Takings.
Calculated: Month is `Format([Paid On], "yyyy-mm")`; From Patients is `Round(IIf([Method] = "Insurer", 0, [Amount]), 2)`; From Insurers is `Round(IIf([Method] = "Insurer", [Amount], 0), 2)`.

#### Takings by practitioner

Money received in the last twelve months for each practitioner's appointments, and how many payments it came in.

From Payments, Invoices, Appointments, Practitioners. Shows Practitioner, Role, Payments, Takings.
Only where `Days Ago Between 0 And 365`.

#### Fees by service

For each service, how many visits were invoiced in the last twelve months and the total charged. It shows which services earn the most.

From Invoices, Appointments, Services. Shows Service, Kind, Visits, Invoiced.
Only where `For = Visit` and `Days Ago Between 0 And 365`.

#### Busy days

Appointments attended or missed in the last year by day of the week, with the hours they took. Use it to plan staff hours.

From Appointments, Services. Shows Day No, Day, Appointments, Hours Booked.
Calculated: Day No is `DateDiff("d", #2024-01-01#, Format([Start], "yyyy-mm-dd")) Mod 7 + 1`; Day is `Mid("MonTueWedThuFriSatSun", (DateDiff("d", #2024-01-01#, Format([Start], "yyyy-mm-dd")) Mod 7) * 3 + 1, 3)`; Hours Booked is `[Minutes] / 60`.
Only where `Status In ("Attended", "No-show")` and `Days Ago Between 1 And 365`.

#### Busy hours

The same appointments by the hour they started, which shows the busy mornings and late afternoons.

From Appointments, Services. Shows Hour, Appointments, Hours Booked.
Calculated: Hour is `Mid([Start], 12, 2) & ":00"`; Hours Booked is `[Minutes] / 60`.
Only where `Status In ("Attended", "No-show")` and `Days Ago Between 1 And 365`.

#### Recalls due

Open recalls falling due in the next 30 days, with contact details and whether a letter has gone out or the patient has agreed to letters at all.

From Recalls, Patients, Services. Shows Due Date, Days Left, Patient, Phone, E-mail, Recall For, Reason, Letters Allowed, Sent On.
Calculated: Days Left is `DateDiff("d", Date(), [Due Date])`; Patient is `[Patients].[First Name] & " " & [Patients].[Last Name]`.
Only where `Days Left Between 0 And 30` and `Booked On Is Null`.

#### Recall letters to send

Recalls due in the next 30 days for patients who agree to letters and have not been sent one yet, with their address. It feeds the recall letters report.

From Recalls, Patients, Services. Shows Recall For, Patient, Address, City, Postcode, Due Date, Days Left.
Calculated: Patient is `[Patients].[First Name] & " " & [Patients].[Last Name]`; Days Left is `DateDiff("d", Date(), [Due Date])`.
Only where `Days Left Between 0 And 30` and `Reminders = Yes` and `Sent On Is Null` and `Booked On Is Null`.

#### New patients by month

How many patients registered each month over the last year, and how many of them are insured.

From Patients, Insurers. Shows Month, New Patients, Insured.
Calculated: Month is `Format([Registered], "yyyy-mm")`; Insured is `IIf([Kind] = "Insurer", 1, 0)`.
Only where `Days Ago Between 0 And 365`.

#### Not seen for six months

Patients whose last attended appointment was more than six months ago, with the date of that visit. Use it for a "we have not seen you for a while" letter or to tidy the list.

From Patients, Appointments. Shows Patient, Patient No, Phone, Last Visit, Days Since.
Calculated: Patient is `[Patients].[First Name] & " " & [Patients].[Last Name]`; Last Visit is `Format([Start], "yyyy-mm-dd")`; Days Since is `DateDiff("d", Format([Start], "yyyy-mm-dd"), Date())`.
Only where `Days Since > 180` and `Status = Attended`.

#### Patient list

All patients by family name with their date of birth, contact details, insurer and policy reference.

From Patients, Insurers. Shows Last Name, First Name, Patient No, Date of Birth, Phone, E-mail, Insurer, Policy Reference, Reminders, Registered.

#### Visit history

Every visit note by patient, newest first, with the practitioner, the service, the reason and the plan.

From Visit Notes, Appointments, Patients, Practitioners, Services. Shows Patient, Patient No, Seen, Practitioner, Service, Reason, Plan.
Calculated: Patient is `[Patients].[First Name] & " " & [Patients].[Last Name]`; Seen is `Format([Start], "yyyy-mm-dd")`.

### Reports

#### Day sheets for the week ahead

A page for each practitioner and day of the week ahead, listing each appointment's time, patient, phone, service and who pays, with the number of patients and minutes booked. Printed each morning for the practitioner and the desk.

From the query Day sheet, grouped by Sheet: Time, Patient, Patient No, Phone, Service, Minutes, Paid By, with count of patient, total of minutes.

#### Invoices due

Unpaid invoices grouped by who owes them, each insurer and each patient, with the balance and days overdue and a total for each. For whoever chases payments.

From the query Unpaid invoices, grouped by Owed By: Invoice, Patient, For, Invoice Date, Due Date, Days Overdue, Balance, with total of balance.

#### Takings by month

A month by month list of money received from patients and from insurers, with totals. For the owners and the accountant.

From the query Takings by month: Month, From Patients, From Insurers, Takings, with total of from patients, total of from insurers, total of takings.

#### Takings by practitioner, last 12 months

Money received in the last twelve months for each practitioner, with a total. Useful when practitioners are paid a share of their takings.

From the query Takings by practitioner: Practitioner, Role, Payments, Takings, with total of takings.

#### Recall letters to send

The names and addresses of patients to send a recall letter to, grouped by what they are recalled for. Print it, or use it for a mail merge.

From the query Recall letters to send, grouped by Recall For: Patient, Address, City, Postcode, Due Date, with count of patient.

#### Patient list

Patients grouped by insurer, with patient number, date of birth, phone and policy reference, and a count for each insurer.

From the query Patient list, grouped by Insurer: Last Name, First Name, Patient No, Date of Birth, Phone, Policy Reference, with count of patient no.

### Starting empty

Start with Insurers: add a Self-pay record first (Kind Self-pay), then each insurer you accept with its payment days. Then add your Practitioners with their hours, then your Services with their minutes and prices. Once those are in, register Patients, pointing each at an insurer or at Self-pay, and book Appointments. Write a Visit Note after each attended appointment, raise an Invoice for it, and record Payments as they arrive, updating the invoice's Amount Paid and Paid On. Add Recalls as visits call for them.

### Making it yours

- For a dental surgery, rename Services to Treatments and add check-ups, scale and polish, fillings and crowns with their minutes and fees, then add a Tooth field to Visit Notes.
- For a vet, add a Pets table pointing at the owner in Patients, with species and breed, and point Appointments at the pet instead of the owner.
- Add a Room field to Appointments, or a Rooms table, if the treatment rooms are shared and you need to see which is free.
- Add a Referred By field to Patients and a query that counts new patients by where they heard of you.
- If you take deposits for a course of treatment, add a Deposit For field to Payments and a status of Deposit paid to Invoices.
- Add an Excess field to Patients for insurers that leave the patient to pay the first part of a claim, and raise a second invoice to the patient for it.
- Add a Consent Signed date to Patients if your practice asks for written consent before treatment, and a query listing patients booked this week who have none.

## Marigold Hair and Beauty

**Salon.** Clients, the appointment book, services, commission, retail stock and gift vouchers for a salon.

For a salon, barber or beauty room that takes bookings, pays its team commission and sells a little retail over the desk. Marigold Hair and Beauty is an invented high-street salon with four stylists, a beauty therapist and a junior. It opens Tuesday to Saturday, with a late night until eight on Thursday.

The records cover the last thirteen months of the appointment book and the next six weeks. Clients come back on their own rhythm, a gents cut every few weeks and a root tint every six or so, and some stop coming. One stylist left about five months ago and most of her clients moved to the stylist who replaced her. December is the busiest month and January the quietest. A few clients cancel late or do not turn up, and the salon keeps their deposit.

Every service done is kept with the price charged, which depends on whether a junior, a stylist or a senior stylist did it, and new clients have 20 per cent off their first visit. Colour work and lash lifts need a patch test in the six months before. Retail and back-bar products are counted in and out as they are sold or used, and a weekly order goes in on Tuesday for whatever has run low. Two lines have been on back order with the supplier for the last few weeks. Gift vouchers are sold, mostly in December and before Mothering Sunday, and redeemed against bookings.

It suits, with small changes: A barber shop; A nail bar; A beauty or brow and lash studio; A massage or spa room; A dog groomer; A mobile hairdresser.

### Tables

#### Suppliers

The wholesalers the salon buys its products from. The sample has 4 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Name | Short text, required | The supplier's trading name. |
| Contact | Short text | The person the salon deals with there. |
| Phone | Short text | Their order line. |
| E-mail | Short text | Where orders and queries go. |
| Account Number | Short text | The salon's account number with the supplier, quoted when ordering. |
| Lead Days | Whole number | How many days an order usually takes to arrive. |

#### Staff

One record per member of the team, past and present, with the commission each earns. The sample has 7 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| First Name | Short text | The name clients know them by. |
| Last Name | Short text, required | Their family name. |
| Role | Short text | Their job in the salon, such as Senior Stylist or Beauty Therapist. |
| Level | Short text | Junior, Stylist, Senior or Therapist: which price column of the service menu their clients pay. |
| Phone | Short text | Their mobile, for rota changes. |
| E-mail | Short text | Their salon e-mail address. |
| Works | Short text | The days they normally work. |
| Started | Date | The day they joined. |
| Left | Date | The day they left, empty while they work here. |
| Service Commission | Number | The share of their service takings paid as commission, 0.12 for 12 per cent. The owner takes none. |
| Retail Commission | Number | The share of the retail they sell paid as commission. |
| Active | Yes/No | Yes while they work here; No keeps their history after they leave. |

#### Services

The service menu: every treatment the salon offers, how long it takes and what it costs at each level. The sample has 26 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Name | Short text, required, unique | The service as it appears on the menu. |
| Category | Short text | The heading it sits under: Cut and Finish, Colour, Treatments, Nails, Brows and Lashes, Skin and Body. |
| Minutes | Whole number | The time to book for it, processing included. |
| Junior Price | Currency | The price with the junior. Empty where the junior does not do it. |
| Stylist Price | Currency | The price with a stylist, and the price for beauty services. |
| Senior Price | Currency | The price with a senior stylist or the owner. |
| Patch Test | Yes/No | Yes where a skin test is needed in the six months before, such as colour and lash lifts. |
| Active | Yes/No | No takes a service off the menu without losing its history. |

#### Clients

One record per client, with what the salon needs to know before they sit down. The sample has 1,292 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| First Name | Short text | Their first name. |
| Last Name | Short text, required | Their family name. |
| Phone | Short text | Their mobile, for reminders and running-late calls. |
| E-mail | Short text | Their e-mail address, for booking confirmations. |
| Heard Of Us | Short text | How they found the salon, to see which kind of advertising brings people in. |
| Client Since | Date | The day they first booked. |
| Patch Test Date | Date | The day of their last skin test for colour or tint. |
| Allergies | Short text | Anything to avoid, such as latex gloves or a fragrance. |
| Colour Notes | Short text | The colour they have and how it is mixed, so anyone can repeat it. |
| Marketing Consent | Yes/No | Yes if they agreed to hear about offers. Only these clients go on the win-back list. |
| Notes | Long text | Anything else worth knowing. |

#### Bookings

The appointment book: one record per appointment, past and to come. The sample has 6,655 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Client | Points at Clients, required | Who the appointment is for. |
| Staff | Points at Staff, required | Who it is booked with. |
| Start | Date and time | The day and time it starts. |
| Minutes | Whole number | How long it is booked for, the services added together. |
| Status | Short text | Booked, Completed, Cancelled or No-show. |
| Booked On | Date | The day the booking was made. |
| Booked Via | Short text | Phone, Online or In salon (rebooked at the desk). |
| Deposit | Currency | Taken when booking new clients and long appointments. |
| Deposit Refunded | Yes/No | Yes when a cancellation came with two days' notice or more and the deposit went back. |
| Cancelled On | Date | The day it was cancelled, for a cancelled booking. |
| Paid By | Short text | Card, Cash or Gift voucher, once the appointment is done. |
| Notes | Long text | Anything to know for this visit. |

#### Booking Services

The services in each booking, with the price charged for each. The sample has 9,229 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Booking | Points at Bookings, required | The appointment it belongs to. |
| Service | Points at Services, required | The service from the menu. |
| Price | Currency | What was charged: the price for the stylist's level, less the new-client offer on a first visit. |

#### Products

Everything the salon stocks: retail products sold to clients and back-bar products used in services. The sample has 38 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Name | Short text, required | The product, described plainly. |
| Kind | Short text | Retail (sold) or Back bar (used at the basin and in treatments). |
| Category | Short text | Shampoo, Colour, Nails and so on. |
| Supplier | Points at Suppliers | Who it is ordered from. |
| Size | Short text | The size of one unit. |
| Cost Price | Currency | What one unit costs the salon. |
| Sale Price | Currency | The retail price; empty for back-bar products. |
| In Stock | Whole number | Units on the shelf now. |
| Reorder Level | Whole number | At or below this many, it goes on the next order. |
| Reorder Quantity | Whole number | How many to order at a time. |
| Active | Yes/No | No for a product no longer stocked. |

#### Product Sales

Retail sold at the desk, one record per product per sale. The sample has 1,134 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Date | Date | The day of the sale. |
| Staff | Points at Staff | Who sold it, for retail commission. |
| Client | Points at Clients | Who bought it. |
| Product | Points at Products, required | What was sold. |
| Quantity | Whole number | How many. |
| Unit Price | Currency | The price of one, as charged. |

#### Stock Deliveries

What came in from suppliers, one record per product per delivery. The sample has 295 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Supplier | Points at Suppliers, required | Who delivered it. |
| Product | Points at Products, required | What came in. |
| Received | Date | The day it arrived. |
| Quantity | Whole number | How many units. |
| Unit Cost | Currency | What each cost on the invoice. |
| Invoice Number | Short text | The supplier's invoice, to match against the bill. |

#### Gift Vouchers

Vouchers sold, and when and where each was used. The sample has 141 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Code | Short text, required, unique | The number printed on the voucher. |
| Sold On | Date | The day it was bought. |
| Bought By | Short text | Who paid for it. |
| For | Short text | Who it is for. |
| Value | Currency | What it is worth. |
| Expires | Date | A year after it was sold. |
| Redeemed On | Date | The day it was used, empty until then. |
| Booking | Points at Bookings | The appointment it paid for. |

### Relationships

- Bookings.Client points at Clients.ID
- Bookings.Staff points at Staff.ID
- Booking Services.Booking points at Bookings.ID
- Booking Services.Service points at Services.ID
- Products.Supplier points at Suppliers.ID
- Product Sales.Staff points at Staff.ID
- Product Sales.Client points at Clients.ID
- Product Sales.Product points at Products.ID
- Stock Deliveries.Supplier points at Suppliers.ID
- Stock Deliveries.Product points at Products.ID
- Gift Vouchers.Booking points at Bookings.ID

### Queries

#### Bookings next 7 days

Every booked appointment from today to a week ahead, by time, with the client's phone number. Use it to send reminders and plan the week.

From Bookings, Clients, Staff. Shows When, Stylist, Client, Phone, Minutes, Deposit, Booked Via.
Calculated: When is `Left([Start], 10) & " " & Mid([Start], 12, 5)`; Stylist is `[Staff].[First Name] & " " & [Staff].[Last Name]`; Client is `[Clients].[First Name] & " " & [Clients].[Last Name]`.
Only where `Days Away Between 0 And 7` and `Status = Booked`.

#### Takings by stylist last month

What each member of the team took in services last calendar month, and how many services they did.

From Bookings, Staff, Booking Services. Shows Stylist, Role, Services Done, Takings.
Calculated: Stylist is `[Staff].[First Name] & " " & [Staff].[Last Name]`.
Only where `Status = Completed` and `Months Ago = 1`.

#### Service commission last month

Each service done last month by someone on commission, with the price and the commission it earns. The commission statement prints it.

From Booking Services, Bookings, Staff, Clients, Services. Shows Stylist, Day, Client, Service, Price, Rate, Commission.
Calculated: Stylist is `[Staff].[First Name] & " " & [Staff].[Last Name]`; Day is `Format([Start], "yyyy-mm-dd")`; Client is `[Clients].[First Name] & " " & [Clients].[Last Name]`; Commission is `Round([Price] * [Service Commission], 2)`.
Only where `Service Commission > 0` and `Status = Completed` and `Months Ago = 1`.

#### Retail commission last month

Retail each person sold last month and the commission due on it.

From Product Sales, Staff, Products. Shows Stylist, Items Sold, Retail Sales, Rate, Commission.
Calculated: Stylist is `[Staff].[First Name] & " " & [Staff].[Last Name]`; Retail Sales is `Round([Quantity] * [Unit Price], 2)`; Commission is `Round([Quantity] * [Unit Price] * [Retail Commission], 2)`.
Only where `Retail Commission > 0` and `Months Ago = 1`.

#### Service takings by month

Service takings for each month, completed appointments only. Shows the busy and quiet months.

From Bookings, Booking Services. Shows Month, Services Done, Takings.
Calculated: Month is `Format([Start], "yyyy-mm")`.
Only where `Status = Completed`.

#### Retail sales by month

Retail sold each month and the profit on it after cost.

From Product Sales, Products. Shows Month, Items Sold, Sales, Profit.
Calculated: Month is `Format([Date], "yyyy-mm")`; Sales is `Round([Quantity] * [Unit Price], 2)`; Profit is `Round([Quantity] * ([Unit Price] - [Cost Price]), 2)`.

#### Clients to win back

Clients who agreed to hear from the salon, last seen between three months and a year ago, with nothing booked. The ones to send an offer to.

From Clients, Bookings. Shows Client, Phone, E-mail, Last Visit, Visits, Days Since.
Calculated: Client is `[Clients].[First Name] & " " & [Clients].[Last Name]`; Last Visit is `Format([Start], "yyyy-mm-dd")`; Days Since is `DateDiff("d", [Start], Date())`.
Only where `Days Since Between 91 And 365` and `Status In ("Completed", "Booked")` and `Marketing Consent = Yes`.

#### Most booked services

Which services were done most over the last year, and what each took.

From Booking Services, Services, Bookings. Shows Service, Category, Times Done, Takings.
Only where `Status = Completed` and `Days Ago Between 0 And 365`.

#### Products to order

Retail and back-bar products at or below their reorder level, by supplier, with what the order will cost.

From Products, Suppliers. Shows Supplier, Product, Kind, Size, In Stock, Reorder Level, Reorder Quantity, Order Cost.
Calculated: Order Cost is `Round([Reorder Quantity] * [Cost Price], 2)`.
Only where `Above Reorder Level <= 0` and `Active = Yes`.

#### No-shows by client

Clients who have missed an appointment, with their late cancellations and the deposits kept. Use it before taking another booking without a deposit.

From Bookings, Clients. Shows Client, Phone, No-shows, Late Cancellations, Deposits Kept.
Calculated: Client is `[Clients].[First Name] & " " & [Clients].[Last Name]`; No-shows is `IIf([Status] = "No-show", 1, 0)`; Late Cancellations is `IIf([Status] = "Cancelled" And DateDiff("d", [Cancelled On], [Start]) < 2, 1, 0)`; Deposits Kept is `IIf([Deposit Refunded], 0, Nz([Deposit]))`.
Only where `No-shows >= 1` and `Status In ("No-show", "Cancelled")`.

#### Patch tests due

Colour and lash appointments in the next two weeks where the client has had no patch test in the six months before. Call them in for one at least two days ahead.

From Bookings, Clients, Booking Services, Services. Shows When, Client, Phone, Last Patch Test, Service.
Calculated: When is `Left([Start], 10) & " " & Mid([Start], 12, 5)`; Client is `[Clients].[First Name] & " " & [Clients].[Last Name]`.
Only where `Patch Test = Yes` and `Status = Booked` and `Days Away Between 0 And 14` and `Days Since Test > 180`.

#### Vouchers outstanding

Gift vouchers sold and not yet used, soonest to expire first. What the salon still owes in services.

From Gift Vouchers. Shows Code, Sold On, Bought By, For, Value, Expires, Days Left.
Calculated: Days Left is `DateDiff("d", Date(), [Expires])`.
Only where `Days Left >= 0` and `Redeemed On Is Null`.

#### New clients by source

New clients over the last year by how they heard of the salon.

From Clients. Shows Heard Of Us, New Clients.
Only where `Days Ago Between 0 And 365`.

#### Retail best sellers

Retail products over the last year, by the profit they made.

From Product Sales, Products. Shows Product, Category, Sold, Sales, Profit.
Calculated: Sales is `Round([Quantity] * [Unit Price], 2)`; Profit is `Round([Quantity] * ([Unit Price] - [Cost Price]), 2)`.
Only where `Days Ago Between 0 And 365`.

#### Service menu

The services on offer, by category, with the price at each level.

From Services. Shows Category, Service, Minutes, Junior Price, Stylist Price, Senior Price.
Only where `Active = Yes`.

### Reports

#### The column: the next 7 days by stylist

The appointment book for the next seven days, a section for each stylist, with times and phone numbers. Print it for the desk.

From the query Bookings next 7 days, grouped by Stylist: When, Client, Phone, Minutes, Deposit, with total of minutes.

#### Commission statement: last month

Last month's services for each person on commission, with a total of takings and commission each. For the owner and for payroll.

From the query Service commission last month, grouped by Stylist: Day, Client, Service, Price, Commission, with total of price, total of commission.

#### Retail commission: last month

Retail commission due for last month, one line a person.

From the query Retail commission last month: Stylist, Items Sold, Retail Sales, Rate, Commission, with total of retail sales, total of commission.

#### Takings by month

Service takings month by month with a grand total.

From the query Service takings by month: Month, Services Done, Takings, with total of services done, total of takings.

#### Products to order

The order list for the week, a section for each supplier, with the cost of each line.

From the query Products to order, grouped by Supplier: Product, Kind, Size, In Stock, Reorder Level, Reorder Quantity, Order Cost, with total of order cost.

#### Clients to win back

A call or mailing list of lapsed clients who agreed to marketing.

From the query Clients to win back: Client, Phone, E-mail, Last Visit, Visits, Days Since, with count of client.

#### Price list

The service menu, printed by category, with the junior, stylist and senior prices.

From the query Service menu, grouped by Category: Service, Minutes, Junior Price, Stylist Price, Senior Price.

### Starting empty

Start with Suppliers, then Staff with their level and commission rates, then the Services menu with a price for each level. Add your Products next, with their supplier and reorder level. Clients can be added as they book. Then make Bookings and add the services for each in Booking Services, changing the price where you charge something different. Record Product Sales and Stock Deliveries as they happen, and Gift Vouchers when you sell them.

### Making it yours

- A barber can drop the beauty services, keep Junior and Senior prices and add a Walk-in choice to Booked Via.
- A dog groomer can rename Clients to Owners and add a Pets table pointing at it (name, breed, size, temperament), then make each booking point at a pet.
- To pay commission only above a target, add a Target field to Staff and a calculated column in Takings by stylist last month that takes it off the takings.
- Add a Reminder Sent yes/no to Bookings and tick it when the reminder goes, so nobody sends two.
- For a nail bar or brow studio, add a Station or Room field to Bookings to show which chair or room each appointment uses.
- Change the patch test rule from six months to another period by editing 180 in the Patch tests due query.
- Add a Tips field to Bookings and a query that totals tips by stylist each month.

## Ironbridge Repairs

**Repair shop.** Jobs, devices, quotes, parts and payments for a shop that mends phones, tablets, laptops and small appliances.

Ironbridge Repairs is an invented repair shop on a high street. Three technicians work at the bench and one assistant runs the counter, where customers book in a phone with a cracked screen or a coffee machine that has stopped pumping, and come back later to collect it.

Every job is one record. It starts as Booked in, is diagnosed, gets a quote, waits for the customer to approve it, waits for a part if the shelf has none, is repaired and is marked Ready, then Collected. A job can also end as Unrepairable, or Ready with the quote turned down and a small diagnosis charge. The parts and labour each job used are kept beside it, so the charge can be explained line by line, and every payment, deposits included, is recorded against its job.

The records cover about fifteen months up to the day the file is made. Phones come in more in the summer holidays and laptops when the school year starts; kitchen appliances pick up before the year ends. Some jobs are waiting on parts today, some customers have not picked up their devices after weeks, a few repairs have come back under warranty, and the parts shelf has items at or below their reorder level, one of them held up because the supplier has it on back order.

It suits, with small changes: A phone and tablet repair kiosk; A computer support firm that takes machines in; A bike shop workshop, with bikes as the devices; A watch and jewellery repairer; An appliance engineer who also visits homes; A musical instrument or audio repair bench.

### Tables

#### Customers

One record for each person or business that brings work in. Business customers on account are invoiced and pay later; everyone else pays at the counter. The sample has 1,676 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| First Name | Short text | The customer's first name, or the contact's at a business. |
| Last Name | Short text, required | The family name. Needed for every customer. |
| Company | Short text | The business name for customers on account. Empty for members of the public. |
| Phone | Short text | The number the shop rings or texts when the device is ready. |
| E-mail | Short text | Where receipts and reminders go, when the customer gave one. |
| Address | Short text | The street address, used for invoices and for home visits. |
| City | Short text | The town or city. |
| Postcode | Short text | The postcode. |
| Contact By | Short text | How the customer wants to hear that the job is ready: Text message, Phone call or E-mail. |
| On Account | Yes/No | Yes for a business that is invoiced and pays by bank transfer after collection. |
| Joined | Date | The day the customer first brought something in. |

#### Staff

The people at the shop: the technicians who take jobs and the counter assistant. The sample has 4 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| First Name | Short text | First name. |
| Last Name | Short text, required | Family name. |
| Role | Short text | Senior Technician, Technician or Counter Assistant. |
| Skills | Short text | What each person is best at, so the counter books jobs to the right bench. |
| Phone | Short text | Work mobile. |
| E-mail | Short text | Work e-mail address. |
| Started | Date | The day they joined the shop. |
| Active | Yes/No | No once someone has left, so they stop appearing in lists of who can take a job. |

#### Suppliers

The firms parts are bought from. The sample has 4 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Name | Short text, required | The supplier's name. |
| Contact | Short text | The person the shop deals with there. |
| Phone | Short text | Their trade counter number. |
| E-mail | Short text | Where orders are sent. |
| Supplies | Short text | What the shop buys from them. |
| Lead Days | Whole number | How many working days an order usually takes to arrive, used to set the expected date. |
| Our Account | Short text | The shop's account reference with the supplier (an invented reference here). |

#### Parts

Every part the shop fits, whether kept on the shelf or ordered in for each job. The sample has 34 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Code | Short text, unique | The shop's own code for the part, printed on the bin label. |
| Name | Short text, required | What the part is, in plain words. |
| Fits | Short text | The kind of device it is for: Phone, Tablet, Laptop or an appliance. |
| Supplier | Points at Suppliers | Who it is bought from. |
| Cost | Currency | What one costs the shop. |
| Price | Currency | What one is charged at on a job. |
| In Stock | Whole number | How many are on the shelf now, including any already set aside for a job. |
| Reorder Level | Whole number | When stock falls to this number the part is ordered again. |
| Reorder Quantity | Whole number | How many to order at a time. |
| Stocked | Yes/No | Yes for parts kept on the shelf. No for parts ordered in for each job, such as main boards and model-specific keyboards. |
| Bin | Short text | Where it is kept on the parts wall. |

#### Devices

Each device a customer has brought in. One customer can have several, and a device that comes back keeps its record, so its history is in one place. The sample has 2,043 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Customer | Points at Customers, required | Who owns it. |
| Kind | Short text | Phone, Tablet, Laptop, Coffee machine, Vacuum cleaner, Food mixer or Sewing machine. |
| Make | Short text | The maker's name (invented makes in this sample). |
| Model | Short text | The model name or number. |
| Serial | Short text | The serial number, checked when the device is handed back. |
| Colour | Short text | Its colour, which helps find it on the shelf. |

#### Jobs

One record for each repair, from booking in to collection. The record number is the job number on the ticket. The sample has 2,401 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Device | Points at Devices, required | The device being repaired, and through it the customer. |
| Technician | Points at Staff | Who the job is booked to. |
| Booked In | Date | The day the device was handed in. |
| Fault | Short text | The fault as the customer described it at the counter. |
| Left With It | Short text | Anything left with the device, such as a charger, case or water tank. |
| Promised | Date | The day the customer was told it would be ready. |
| Status | Short text | Booked in, Diagnosing, Awaiting approval, Awaiting parts, Repairing, Ready, Collected or Unrepairable. |
| Quote | Currency | The price given for the repair. Fixed-price jobs such as screens are quoted at the counter; others after diagnosis. |
| Quoted On | Date | The day the quote was given. |
| Quote Approved | Yes/No | Yes once the customer has agreed to the quote. |
| Deposit | Currency | Money taken up front when a part has to be ordered in. |
| Outcome | Short text | How the job ended: Repaired, Quote declined or Beyond economic repair. |
| Charge | Currency | What the job finally costs the customer. Nothing for a device that cannot be repaired ("no fix, no fee") or a warranty return. |
| Finished | Date | The day the work was done and the customer could be told. |
| Collected | Date | The day the customer took the device away. |
| Paid | Yes/No | Yes once everything charged has been paid. |
| Warranty Until | Date | The last day of the 90-day warranty on the repair. |
| Warranty Of | Points at Jobs | For a warranty return, the number of the first job it came back from. |
| Notes | Long text | Anything else the counter or the bench needs to know, such as reminders sent. |

#### Part Orders

Parts ordered from suppliers: stock to refill the shelf, or one part for one job. The sample has 762 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Supplier | Points at Suppliers, required | Who the order went to. |
| Part | Points at Parts, required | The part ordered. |
| Job | Points at Jobs | The job the part is for, when it was ordered for one. Empty for shelf stock. |
| Quantity | Whole number | How many were ordered. |
| Unit Cost | Currency | The price paid for each. |
| Ordered On | Date | The day the order was placed. |
| Expected | Date | The day the supplier said it would arrive. |
| Received On | Date | The day it actually arrived. Empty while it is still on its way. |
| Status | Short text | On order or Received. |

#### Job Parts

The parts used on each job, at the price charged. The sample has 1,994 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Job | Points at Jobs, required | The job. |
| Part | Points at Parts, required | The part fitted or waiting to be fitted. |
| Quantity | Whole number | How many. |
| Unit Price | Currency | The price charged for each (nothing on a warranty return). |
| Fitted On | Date | The day it was fitted. Empty while the job is waiting for it. |

#### Labour

The bench time on each job: diagnosis, the repair itself, and warranty work. The sample has 2,885 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Job | Points at Jobs, required | The job. |
| Done On | Date | The day the work was done. |
| Description | Short text | What was done. |
| Hours | Number | How long it took, in hours. |
| Rate | Currency | The hourly rate charged. Nought for warranty work and for checks on devices that could not be repaired. |

#### Payments

Every payment taken, each against its job. The sample has 2,494 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Job | Points at Jobs, required | The job paid for. |
| Paid On | Date | The day the money came in. |
| Amount | Currency | How much. |
| Method | Short text | Card, Cash or Bank transfer. |
| Kind | Short text | Deposit, On collection, or On account for a business paying its invoice. |

### Relationships

- Parts.Supplier points at Suppliers.ID
- Devices.Customer points at Customers.ID
- Jobs.Device points at Devices.ID
- Jobs.Technician points at Staff.ID
- Jobs.Warranty Of points at Jobs.ID
- Part Orders.Supplier points at Suppliers.ID
- Part Orders.Part points at Parts.ID
- Part Orders.Job points at Jobs.ID
- Job Parts.Job points at Jobs.ID
- Job Parts.Part points at Parts.ID
- Labour.Job points at Jobs.ID
- Payments.Job points at Jobs.ID

### Queries

#### Job board

Every job not yet collected, in the order the shop works through them, with who it belongs to, the technician, the promised date and how many days are left (below nought means it is late). Open it first thing in the morning.

From Jobs, Devices, Customers, Staff. Shows Stage, Job, Booked In, Customer, Phone, Device, Fault, Technician, Promised, Days Left.
Calculated: Stage is `IIf([Status] = "Booked in", "1 Booked in", IIf([Status] = "Diagnosing", "2 Diagnosing", IIf([Status] = "Awaiting approval", "3 Awaiting approval", IIf([Status] = "Awaiting parts", "4 Awaiting parts", IIf([Status] = "Repairing", "5 Repairing", IIf([Status] = "Ready", "6 Ready", "7 " & [Status]))))))`; Customer is `[Customers].[First Name] & " " & [Customers].[Last Name]`; Device is `[Make] & " " & [Model]`; Technician is `[Staff].[First Name] & " " & [Staff].[Last Name]`; Days Left is `IIf(Nz([Finished], "") = "", DateDiff("d", Date(), [Promised]), Null)`.
Only where `Collected Is Null`.

#### Jobs by status

How many jobs sit at each stage now, and the oldest booking at each. A quick count for the morning meeting.

From Jobs. Shows Stage, Jobs, Oldest Booked In.
Calculated: Stage is `IIf([Status] = "Booked in", "1 Booked in", IIf([Status] = "Diagnosing", "2 Diagnosing", IIf([Status] = "Awaiting approval", "3 Awaiting approval", IIf([Status] = "Awaiting parts", "4 Awaiting parts", IIf([Status] = "Repairing", "5 Repairing", IIf([Status] = "Ready", "6 Ready", "7 " & [Status]))))))`.
Only where `Collected Is Null`.

#### Awaiting parts

Jobs held up for a part, with the part, the supplier and when it is expected. Use it when a customer rings to ask, or to chase a supplier.

From Part Orders, Parts, Suppliers, Jobs, Devices, Customers. Shows Job, Customer, Phone, Device, Part, Supplier, Ordered On, Expected, Promised.
Calculated: Customer is `[Customers].[First Name] & " " & [Customers].[Last Name]`; Device is `[Make] & " " & [Model]`.
Only where `Received On Is Null` and `Status = Awaiting parts`.

#### Past promised date

Unfinished jobs whose promised date has gone by, latest first, with the customer's number so the counter can ring them with news.

From Jobs, Devices, Customers, Staff. Shows Job, Status, Customer, Phone, Device, Technician, Promised, Days Late.
Calculated: Customer is `[Customers].[First Name] & " " & [Customers].[Last Name]`; Device is `[Make] & " " & [Model]`; Technician is `[Staff].[First Name] & " " & [Staff].[Last Name]`; Days Late is `DateDiff("d", [Promised], Date())`.
Only where `Days Late > 0` and `Finished Is Null`.

#### Ready for collection

Every finished device still on the shelf, longest waiting first, with what is left to pay and how the customer wants to be told. The calling list.

From Jobs, Devices, Customers. Shows Job, Customer, Phone, Contact By, Device, Outcome, Finished, Days Waiting, To Pay.
Calculated: Customer is `[Customers].[First Name] & " " & [Customers].[Last Name]`; Device is `[Make] & " " & [Model]`; Days Waiting is `DateDiff("d", [Finished], Date())`; To Pay is `Round(Nz([Charge]) - Nz([Deposit]), 2)`.
Only where `Finished Is Not Null` and `Collected Is Null`.

#### Uncollected 14 days or more

Devices ready for two weeks or longer and not picked up, with any reminders already sent. Use it to send a final reminder before the shop's storage terms apply.

From Jobs, Devices, Customers. Shows Job, Customer, Phone, E-mail, Device, Finished, Days Waiting, To Pay, Notes.
Calculated: Customer is `[Customers].[First Name] & " " & [Customers].[Last Name]`; Device is `[Make] & " " & [Model]`; Days Waiting is `DateDiff("d", [Finished], Date())`; To Pay is `Round(Nz([Charge]) - Nz([Deposit]), 2)`.
Only where `Days Waiting >= 14` and `Collected Is Null`.

#### Quotes awaiting approval

Quotes given and not yet answered, oldest first, with the price and the phone number to follow up.

From Jobs, Devices, Customers. Shows Job, Quoted On, Days Waiting, Customer, Phone, Device, Fault, Quote.
Calculated: Days Waiting is `DateDiff("d", [Quoted On], Date())`; Customer is `[Customers].[First Name] & " " & [Customers].[Last Name]`; Device is `[Make] & " " & [Model]`.
Only where `Status = Awaiting approval`.

#### Parts to reorder

Shelf parts at or below their reorder level, by supplier, with what a full reorder would cost.

From Parts, Suppliers. Shows Supplier, Code, Part, In Stock, Reorder Level, Reorder Quantity, Cost, Order Value.
Calculated: Order Value is `Round([Reorder Quantity] * [Cost], 2)`.
Only where `Spare <= 0` and `Stocked = Yes`.

#### Parts on order

Everything ordered and not yet arrived, with the expected date and the job it is for when there is one.

From Part Orders, Parts, Suppliers. Shows Supplier, Part, Quantity, Unit Cost, Ordered On, Expected, For Job.
Only where `Received On Is Null`.

#### Takings by month

Money taken each month, deposits and account payments included, and how many payments made it up.

From Payments. Shows Month, Payments, Takings.
Calculated: Month is `Format([Paid On], "yyyy-mm")`.

#### Takings last 30 days

The last thirty days of takings split by payment method, for checking against the card machine and the bank statement.

From Payments. Shows Method, Payments, Takings.
Only where `Days Ago Between 0 And 30`.

#### Technician workload

How many unfinished jobs each technician has, how many are stuck on parts or quotes, and their oldest job.

From Jobs, Staff. Shows Technician, Open Jobs, Awaiting Parts, Awaiting Approval, Oldest Booked In.
Calculated: Technician is `[Staff].[First Name] & " " & [Staff].[Last Name]`; Awaiting Parts is `IIf([Status] = "Awaiting parts", 1, 0)`; Awaiting Approval is `IIf([Status] = "Awaiting approval", 1, 0)`.
Only where `Finished Is Null`.

#### Technician turnaround

For repairs finished in the last year: how many each technician did, the average days from booking in to finished, and what they were charged at.

From Jobs, Staff. Shows Technician, Jobs Repaired, Average Days, Charged.
Calculated: Technician is `[Staff].[First Name] & " " & [Staff].[Last Name]`; Average Days is `DateDiff("d", [Booked In], [Finished])`.
Only where `Outcome = Repaired` and `Days Ago Between 0 And 365`.

#### Work by device kind

Jobs and money by kind of device over the whole period, with the average days each kind takes.

From Jobs, Devices. Shows Kind, Jobs, Charged, Average Days.
Calculated: Average Days is `DateDiff("d", [Booked In], [Finished])`.

#### Warranty returns

Jobs that came back under warranty, with the number of the first job, so repeat faults on a part or a technician's work show up.

From Jobs, Devices, Customers, Staff. Shows Job, Booked In, First Job, Customer, Device, Fault, Technician, Status.
Calculated: Customer is `[Customers].[First Name] & " " & [Customers].[Last Name]`; Device is `[Make] & " " & [Model]`; Technician is `[Staff].[First Name] & " " & [Staff].[Last Name]`.
Only where `Warranty Of Is Not Null`.

#### Job sheets

Every job booked in over the last fourteen days with its full detail: serial, fault, what was left with it, quote and promised date.

From Jobs, Devices, Customers, Staff. Shows Technician, Job, Booked In, Customer, Phone, Device, Serial, Fault, Left With It, Status, Quote, Promised.
Calculated: Technician is `[Staff].[First Name] & " " & [Staff].[Last Name]`; Customer is `[Customers].[First Name] & " " & [Customers].[Last Name]`; Device is `[Make] & " " & [Model]`.
Only where `Days Ago Between 0 And 14`.

#### Parts margin

Each part fitted over the period: how many, what they were charged at, and the margin over cost. Warranty fittings count against the margin.

From Job Parts, Parts. Shows Fits, Part, Fitted, Charged, Margin.
Calculated: Charged is `Round([Quantity] * [Unit Price], 2)`; Margin is `Round([Quantity] * ([Unit Price] - [Cost]), 2)`.
Only where `Fitted On Is Not Null`.

#### Accounts to chase

Business customers on account who have collected work and not yet paid for it, oldest first.

From Jobs, Devices, Customers. Shows Company, Contact, E-mail, Job, Collected, Charge, Days Since Collected.
Calculated: Contact is `[Customers].[First Name] & " " & [Customers].[Last Name]`; Days Since Collected is `DateDiff("d", [Collected], Date())`.
Only where `Collected Is Not Null` and `Paid = No` and `Charge > 0`.

### Reports

#### Job board

The job board printed for the bench, one section per stage, each job with its customer, device, fault, technician and promised date.

From the query Job board, grouped by Stage: Job, Booked In, Customer, Device, Fault, Technician, Promised, Days Left, with count of job.

#### Ready for collection: to phone

A calling list for the counter: each device waiting to be collected, the customer's number, how they like to be contacted and what they owe, longest waiting first.

From the query Ready for collection: Job, Customer, Phone, Contact By, Device, Outcome, Finished, Days Waiting, To Pay, with count of job, total of to pay.

#### Parts to order

The parts to reorder, one section per supplier, with the cost of each order. Take it to the phone or the supplier's website.

From the query Parts to reorder, grouped by Supplier: Code, Part, In Stock, Reorder Level, Reorder Quantity, Cost, Order Value, with total of order value.

#### Takings by month

Monthly takings with a grand total, for the accountant or the owner's month-end look.

From the query Takings by month: Month, Payments, Takings, with total of payments, total of takings.

#### Job sheets: the last two weeks

The last two weeks of jobs, one section per technician, with every detail taken at the counter. Print it to check work in progress against the shelf.

From the query Job sheets, grouped by Technician: Job, Booked In, Customer, Phone, Device, Serial, Fault, Left With It, Status, Quote, Promised, with count of job.

#### Technician turnaround

Each technician's finished repairs over the last year, their average turnaround and the value of their work.

From the query Technician turnaround: Technician, Jobs Repaired, Average Days, Charged, with total of jobs repaired, total of charged.

### Starting empty

Start with Staff, then Suppliers, then Parts (each part pointing at its supplier, with Stocked set to No for parts you only order in). When a customer comes to the counter, add them to Customers, add their device to Devices, then open a Job for that device with the fault, the technician and the promised date. As the job moves on, change its Status and fill in the quote, the parts used (Job Parts) and the time spent (Labour). Record each payment in Payments as it is taken, and put Finished and Collected on the job when they happen.

### Making it yours

- For a bike workshop, use Devices for bikes (Kind as Road, Mountain, Hybrid, E-bike; Serial as the frame number) and add a Booked For date for the day the customer wants it back.
- For an engineer who visits homes, add Visit Date and Visit Address to Jobs and a Call-out Fee to the charge, and make a query of visits in the next 7 days ordered by postcode.
- Add a Courier Reference to Part Orders, and a query of orders past their Expected date, to chase late suppliers.
- Add a status such as Sent to maker for repairs sent away, and add it to the Stage calculation in the Job board query so it sorts in the right place.
- Add a Disposal Date to Jobs for devices never collected, and a query of devices uncollected for 90 days or more, if your terms say you may then recycle or sell them.
- Give Staff an hourly rate and add a query that sets Labour hours against it, to see what each technician earns the shop.
- Add Passcode Given (Yes/No) and Data Backed Up (Yes/No) to Jobs, so the bench knows what it can test before starting.

## Bramble Studio

**Consultancy or agency.** Clients, projects, time sheets, expenses and invoices for a small studio or consultancy.

For a small firm that sells its time: a design studio, a consultancy, a bookkeeper or anyone who quotes for work, logs hours against it and invoices for it. It answers the questions such a firm asks each week: what time has not been billed yet, which projects are eating their budget, who is owed what and for how long.

Bramble Studio is invented. Six people work there: two partners, a senior designer, a designer, a developer and a project manager. They make brand identities, websites, annual reports, packaging and campaigns for about two dozen small and mid-sized clients, most in the UK and a few in Ireland and the US. Some work is for a fixed fee paid in stages; some is by the hour, invoiced at the start of each month for the month before, including three retainers that renew.

The records cover the last sixteen months, counted back from the day the file is made. Proposals come in every few days and about two in three are won, fewer when the studio is already busy. Every hour is logged to a task, including the studio's own time on admin and new business, so the billable share of each person's week shows. Some projects overrun, a few are paused by the client for a while, some invoices are paid late or in two parts, and the studio is quiet in August and shut over Christmas.

It suits, with small changes: A freelance designer or developer; A marketing or branding agency; An architect's practice; IT or management consultants; An accountant or bookkeeper who bills by the hour; A small law practice; A translation or copywriting business; A video or photography studio.

### Tables

#### Clients

One record per client the studio works for or has quoted for. The studio itself is kept here too, as the client of its own internal projects. The sample has 26 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Name | Short text, required | The client's business name, as it goes on an invoice. |
| Sector | Short text | The line of business the client is in, such as Charity or Retail; Internal for the studio itself. |
| Address | Short text | The street address invoices are sent to. |
| City | Short text | The town or city of the address. |
| Postcode | Short text | The postcode or ZIP of the address. |
| Country | Short text | The country of the address. |
| Phone | Short text | The client's main phone number. |
| E-mail | Short text | Where invoices are sent, usually an accounts address. |
| Payment Days | Whole number | How many days after the invoice date payment is due: the client's terms. |
| VAT Rate | Number | The VAT charged on this client's invoices as a fraction: 0.2 for UK clients, 0 for clients abroad. |
| Client Since | Date | When the client first came to the studio. |
| Notes | Long text | Anything worth knowing, such as a client who pays late. |

#### Contacts

The people the studio deals with at each client. A client can have several. The sample has 49 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Client | Points at Clients, required | The client the person works for. |
| First Name | Short text | Their first name. |
| Last Name | Short text, required | Their family name. |
| Job Title | Short text | What they do there, so you know who signs off work and who pays. |
| E-mail | Short text | Their e-mail address. |
| Phone | Short text | Their direct phone number. |
| Main Contact | Yes/No | Yes for the person to speak to first about the work. |

#### Staff

Everyone who logs time. What each hour costs and what it is charged at turn hours into money in the queries. The sample has 6 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| First Name | Short text | Their first name. |
| Last Name | Short text, required | Their family name. |
| Role | Short text | Their job in the studio. |
| E-mail | Short text | Their work e-mail address. |
| Phone | Short text | Their work mobile. |
| Hourly Cost | Currency | What an hour of this person costs the studio, with pay, pension and a share of overheads. |
| Charge-out Rate | Currency | What an hour of this person is charged to a client on hourly work. |
| Hours per Week | Number | Their contracted hours, for checking time sheets against. |
| Started | Date | The day they joined. |
| Active | Yes/No | No once they have left; their time stays in the records. |

#### Projects

One record per piece of work quoted for, from proposal to completion or loss. The studio's own time (admin, new business, training) is kept on three internal projects. The sample has 116 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Code | Short text, unique | A short reference for the project, used on time sheets and in conversation. |
| Name | Short text, required | What the work is, as the client knows it. |
| Client | Points at Clients, required | Who the work is for. |
| Kind | Short text | Fixed fee, Hourly or Internal. It decides how the work is invoiced. |
| Work Type | Short text | The kind of job, such as Brand identity or Website design and build, for comparing like with like. |
| Budget Hours | Number | The hours the fee or estimate was worked out from. |
| Fee | Currency | For fixed-fee work, the agreed fee. For hourly work, the estimate quoted; once the project is completed it becomes what the billable time came to. |
| Proposed On | Date | When the proposal went to the client. |
| Start Date | Date | When the work starts, or could start if the proposal is won. |
| Due Date | Date | When the work is due to be finished. A pause moves it on. |
| Finished On | Date | The day the last of the work was done. |
| Status | Short text | Proposal, Active, On hold, Completed or Lost. |
| Notes | Long text | Why a proposal was lost, why work is paused, and the like. |

#### Tasks

The parts a project is broken into, each with the hours estimated for it, so you can see which part ran over. The sample has 565 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Project | Points at Projects, required | The project the task belongs to. |
| Name | Short text, required | What the task is, such as Concepts or Front-end build. |
| Role | Short text | Who does it: Partner, Design, Development or Project management. |
| Estimated Hours | Number | The hours allowed for it in the budget. |
| Done | Yes/No | Yes when the task is finished. |

#### Time Entries

Every stretch of time someone worked, logged to a project and one of its tasks. The heart of the database. The sample has 4,395 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Date | Date, required | The day the work was done. |
| Staff | Points at Staff, required | Who did it. |
| Project | Points at Projects, required | The project it was for. |
| Task | Points at Tasks | The task of that project it was for. |
| Hours | Number | How long, in hours and quarters of an hour. |
| Billable | Yes/No | Yes for client work; No for the studio's own time and for corrections done at the studio's cost. |
| Note | Short text | A word on what was done. |
| Invoice | Points at Invoices | The invoice the hours were billed on. Empty means not billed yet. On fixed-fee work it is the stage invoice that covered them. |

#### Expenses

Money spent on a project: licences, printing, travel, hosting. Rechargeable ones are passed on to the client at cost. The sample has 145 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Date | Date | When the cost was run up. |
| Project | Points at Projects, required | The project it was for. |
| Staff | Points at Staff | Who spent it. |
| Category | Short text | The kind of cost, such as Printing or Travel. |
| Description | Short text | What it was. |
| Amount | Currency | What it cost, before VAT. |
| Rechargeable | Yes/No | Yes if the client pays for it on their next invoice. |
| Invoice | Points at Invoices | The invoice it was recharged on. Empty means not recharged yet. |

#### Invoices

Every invoice sent, with what was paid on it so far. The sample has 168 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Number | Short text, unique | The invoice number, as printed. |
| Project | Points at Projects, required | The project billed; the client comes from the project. |
| Invoice Date | Date | The date on the invoice. |
| Due Date | Date | When payment is due, from the client's payment days. |
| Net | Currency | The total of the lines before VAT. |
| VAT | Currency | The VAT charged at the client's rate. |
| Total | Currency | Net plus VAT: what the client owes. |
| Paid | Currency | What has been received against it so far, from Payments. |
| Status | Short text | Sent, Part paid or Paid. |

#### Invoice Lines

The lines of each invoice: hours by role on hourly work, a stage of the fee on fixed-fee work, and expenses passed on. The sample has 333 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Invoice | Points at Invoices, required | The invoice the line is on. |
| Description | Short text | What the line says, as printed. |
| Kind | Short text | Time, Fee or Expense. |
| Quantity | Number | Hours for a Time line; 1 for a fee stage or an expense. |
| Unit Price | Currency | The rate per hour, or the amount of the stage or expense. |

#### Payments

Money received from clients. An invoice can be paid in more than one payment. The sample has 166 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Invoice | Points at Invoices, required | The invoice the money was for. |
| Received On | Date | The day it arrived. |
| Amount | Currency | How much arrived. |
| Method | Short text | Bank transfer or Card. |
| Reference | Short text | The reference the client put on the payment, usually the invoice number. |

### Relationships

- Contacts.Client points at Clients.ID
- Projects.Client points at Clients.ID
- Tasks.Project points at Projects.ID
- Time Entries.Staff points at Staff.ID
- Time Entries.Project points at Projects.ID
- Time Entries.Task points at Tasks.ID
- Time Entries.Invoice points at Invoices.ID
- Expenses.Project points at Projects.ID
- Expenses.Staff points at Staff.ID
- Expenses.Invoice points at Invoices.ID
- Invoices.Project points at Projects.ID
- Invoice Lines.Invoice points at Invoices.ID
- Payments.Invoice points at Invoices.ID

### Queries

#### Unbilled time

Billable hours on hourly projects not yet on an invoice, by client, project and person, with what they are worth at charge-out rates. Run it before the monthly invoices go out.

From Time Entries, Projects, Clients, Staff. Shows Client, Project, Person, Earliest, Hours, Value.
Calculated: Person is `[First Name] & " " & [Last Name]`; Value is `Round([Hours] * [Charge-out Rate], 2)`.
Only where `Kind = Hourly` and `Billable = Yes` and `Invoice Is Null`.

#### Unbilled expenses

Rechargeable costs not yet passed on to the client, oldest first, so nothing is forgotten on the next invoice.

From Expenses, Projects, Clients. Shows Client, Project, Date, Category, Description, Amount.
Only where `Rechargeable = Yes` and `Invoice Is Null`.

#### Project budgets

Each project in hand: its budget hours against the hours used so far, and its fee against the time used at charge-out rates and at cost. A project whose time at rates is near its fee needs a word with the client.

From Time Entries, Projects, Clients, Staff. Shows Client, Code, Project, Kind, Due Date, Budget Hours, Hours Used, Fee, Time at Rates, Cost of Time.
Calculated: Time at Rates is `Round([Hours] * [Charge-out Rate], 2)`; Cost of Time is `Round([Hours] * [Hourly Cost], 2)`.
Only where `Kind <> "Internal"` and `Status In ("Active", "On hold")`.

#### Task hours against estimate

For each project in hand, each task's estimated hours against the hours logged to it, to see which part of the work is running over.

From Time Entries, Tasks, Projects. Shows Code, Project, Task, Role, Estimated Hours, Hours Used.
Only where `Status In ("Active", "On hold")` and `Kind <> "Internal"`.

#### Hours by person, last 30 days

Each person's hours over the last 30 days, how many of them were billable and what those were worth. Use it to see who is stretched and who has room.

From Time Entries, Staff. Shows Person, Role, Hours, Billable Hours, Billable Value.
Calculated: Person is `[First Name] & " " & [Last Name]`; Billable Hours is `IIf([Billable], [Hours], 0)`; Billable Value is `Round(IIf([Billable], [Hours] * [Charge-out Rate], 0), 2)`.
Only where `Days Ago Between 0 And 30`.

#### Time sheet

Every entry of the last 14 days, by person and day, with the client, project and task. For checking time sheets before invoicing.

From Time Entries, Tasks, Projects, Clients, Staff. Shows Person, Date, Client, Project, Task, Hours, Billable, Note.
Calculated: Person is `[First Name] & " " & [Last Name]`.
Only where `Days Ago Between 0 And 13`.

#### Overdue invoices

Invoices past their due date with money still owed, the longest overdue first, with the client's phone number for chasing.

From Invoices, Projects, Clients. Shows Invoice, Client, Project, Invoice Date, Due Date, Total, Paid, Outstanding, Days Overdue, Client Phone.
Calculated: Outstanding is `Round([Total] - [Paid], 2)`; Days Overdue is `DateDiff("d", [Invoices].[Due Date], Date())`.
Only where `Outstanding > 0` and `Days Overdue > 0`.

#### Unpaid invoices

Every invoice with money still owed, by client, with the days it is overdue (below nought when it is not due yet).

From Invoices, Projects, Clients. Shows Client, Invoice, Project, Invoice Date, Due Date, Total, Outstanding, Days Overdue.
Calculated: Outstanding is `Round([Total] - [Paid], 2)`; Days Overdue is `DateDiff("d", [Invoices].[Due Date], Date())`.
Only where `Outstanding > 0`.

#### Revenue by client

Each client's invoiced fees before VAT, last year and this year, the biggest this year first. Shows who the studio depends on.

From Invoices, Projects, Clients. Shows Client, Sector, Last Year, This Year, Invoices.
Calculated: Last Year is `Round(IIf(Year([Invoice Date]) = Year(Date()) - 1, [Net], 0), 2)`; This Year is `Round(IIf(Year([Invoice Date]) = Year(Date()), [Net], 0), 2)`.

#### Revenue by month

What was invoiced each month, net, VAT and total, for seeing the busy and quiet months.

From Invoices. Shows Month, Invoices, Net, VAT, Total.
Calculated: Month is `Format([Invoice Date], "yyyy-mm")`.

#### Income by kind

Invoiced amounts each year split into fixed-fee stages, hourly time and recharged expenses.

From Invoice Lines, Invoices. Shows Year, Kind, Lines, Amount.
Calculated: Year is `Format([Invoice Date], "yyyy")`; Amount is `Round([Quantity] * [Unit Price], 2)`.

#### Cash received by month

The money that actually came in each month, which is what pays the wages.

From Payments. Shows Month, Payments, Received.
Calculated: Month is `Format([Received On], "yyyy-mm")`.

#### Project profit

Completed projects, newest first: the fee earned against the cost of the time spent, and the time at charge-out rates. Fee less cost of time is what the project made towards the studio's other costs.

From Time Entries, Projects, Clients, Staff. Shows Finished On, Code, Project, Client, Kind, Fee, Hours, Time at Rates, Cost of Time.
Calculated: Time at Rates is `Round([Hours] * [Charge-out Rate], 2)`; Cost of Time is `Round([Hours] * [Hourly Cost], 2)`.
Only where `Kind <> "Internal"` and `Status = Completed`.

#### Projects due

Projects in hand due within the next 30 days, and any already late (Days Left below nought), soonest first.

From Projects, Clients. Shows Due Date, Days Left, Code, Project, Client, Kind, Start Date, Status.
Calculated: Days Left is `DateDiff("d", Date(), [Due Date])`.
Only where `Days Left <= 30` and `Status In ("Active", "On hold")`.

#### Proposals waiting

Proposals sent and not yet answered, the oldest first, with how long each has waited. For a follow-up call.

From Projects, Clients. Shows Proposed On, Days Waiting, Code, Project, Client, Kind, Budget Hours, Fee, Could Start.
Calculated: Days Waiting is `DateDiff("d", [Proposed On], Date())`.
Only where `Status = Proposal`.

#### Recent invoice lines

The lines of every invoice sent in the last 45 days, for checking what went out.

From Invoice Lines, Invoices, Projects, Clients. Shows Invoice, Invoice Date, Client, Project, Description, Quantity, Unit Price, Amount.
Calculated: Amount is `Round([Quantity] * [Unit Price], 2)`.
Only where `Days Ago Between 0 And 45`.

#### Client contacts

The people at each client, the main contact first, with their e-mail and phone.

From Contacts, Clients. Shows Client, Main Contact, First Name, Last Name, Job Title, E-mail, Phone.

### Reports

#### Time sheet, last 14 days

The last 14 days of time, one section per person with their total hours. For each person to check, and for the project manager before invoicing.

From the query Time sheet, grouped by Person: Date, Client, Project, Task, Hours, Note, with total of hours.

#### Work in progress to invoice

Unbilled hourly time by client, with hours and value, and a total for the studio: what the next round of invoices will bring in.

From the query Unbilled time, grouped by Client: Project, Person, Earliest, Hours, Value, with total of hours, total of value.

#### Invoices due

Money owed by each client, invoice by invoice, with the days overdue and the total outstanding. For the partner who chases payment.

From the query Unpaid invoices, grouped by Client: Invoice, Project, Invoice Date, Due Date, Days Overdue, Total, Outstanding, with total of outstanding.

#### Project budget report

Projects in hand by client, with budget hours, hours used, fee, time at rates and cost of time. For the weekly project meeting.

From the query Project budgets, grouped by Client: Code, Project, Kind, Due Date, Budget Hours, Hours Used, Fee, Time at Rates, Cost of Time, with total of hours used, total of fee, total of time at rates, total of cost of time.

#### Revenue by client

One line per client with last year's and this year's invoiced fees, and the totals.

From the query Revenue by client: Client, Sector, Invoices, Last Year, This Year, with total of last year, total of this year.

#### Completed projects: fee against cost of time

Completed projects with their fee and the cost of the time they took, and the totals for the period.

From the query Project profit: Finished On, Code, Project, Client, Kind, Hours, Fee, Cost of Time, with total of fee, total of cost of time.

#### Invoices issued, last 45 days

Each invoice of the last 45 days with its lines and its total before VAT.

From the query Recent invoice lines, grouped by Invoice: Invoice Date, Client, Description, Quantity, Unit Price, Amount, with total of amount.

### Starting empty

Start with Staff: each person with their hourly cost and charge-out rate. Then Clients and their Contacts, adding your own firm as a client if you want to record internal time. Add a project for each piece of work, with its kind, budget hours and fee, and break it into Tasks. From then on, log Time Entries and Expenses as the work goes. When you invoice, add the Invoice and its Invoice Lines, and put the invoice in the Invoice field of the time entries and expenses it covers, so they drop out of the unbilled queries. Record each Payment when the money arrives and bring the invoice's Paid and Status up to date.

### Making it yours

- If rates differ by client, add a Rate field to Projects and use it instead of the person's charge-out rate in the Unbilled time query.
- A law or accounting practice can rename Projects to Matters and Tasks to Stages, and add a Matter Type and a Fee Earner field.
- Add a Purchase Order field to Projects for clients who will not pay an invoice without their own order number on it.
- Add a Leave table, or an internal Holiday project, so time off shows on the time sheet and the hours per week add up.
- Add a Status of Invoiced to fixed-fee projects whose last stage has gone out but whose work is not signed off yet.
- Add a Mileage expense category with a Miles field, and a calculated column at your mileage rate in the Unbilled expenses query.
- Make a query of Time Entries by Work Type to see which kinds of job overrun most, before quoting for the next one.

## Little Fig Cafe

**Cafe or restaurant.** Menu, recipes and food costs, till orders, rota, deliveries, waste and the nightly cash-up for a small cafe.

For the owner or manager of a small cafe who wants to see what each dish costs to make, what sells and when, whether the wages are in hand against the takings, what to order from each supplier and what is being thrown away.

Little Fig Cafe is an invented all-day cafe with about thirty seats and a takeaway counter. It opens every day but Christmas Day, from half past seven (half past eight at weekends) until half past four. Breakfast is served until half past eleven and lunch until half past three; coffee, cold drinks and cakes all day. Its head baker makes the cakes on site each morning. Eight people work there, and six invented suppliers deliver coffee, dairy, produce, meat and fish, bread and dry goods, each on its own days.

The records cover the last fourteen months: a cash-up sheet for every day the cafe opened, the rota as it was worked, every delivery and its bill, and what was thrown away. The till's orders are here line by line for the last eight weeks, about six thousand of them; older days are in their cash-up sheets. Trade is busier at weekends, in July and August and on sunny days, and quieter in January. The rota runs two weeks ahead. The records stop at the start of today: this morning's deliveries are not booked in yet and this morning's orders to suppliers are not placed, so the order list shows what to ask for.

Every dish has a recipe in its ingredients' own units, so its food cost comes from the ingredients' current prices. Prices are pounds, as a cafe in the UK would charge them, but nothing in the design depends on the currency.

It suits, with small changes: A coffee shop or kiosk; A small restaurant or bistro; A bakery with a few tables; A food truck or market stall; A sandwich bar or takeaway; A tea room or farm shop cafe.

### Tables

#### Menu Categories

The sections of the menu, such as breakfast or cakes. Enter these first. The sample has 5 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Name | Short text, required, unique | The section as the menu heads it. |
| Served | Short text | When the section is served, such as All day or Until 11:30, printed on the menu. |
| Sort Order | Whole number | The place of the section on the printed menu and the costing sheet: 1 comes first. |

#### Menu Items

One record per dish or drink on the menu, past or present. The sample has 39 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Name | Short text, required, unique | The dish as the menu and the till name it. No two dishes share a name. |
| Category | Points at Menu Categories, required | The menu section it belongs to. |
| Price | Currency | What a customer pays for one, as the menu shows it. |
| Takeaway | Yes/No | Yes when it can be taken away as well as eaten in. |
| Baked Here | Yes/No | Yes for what the cafe bakes itself, such as cakes and pancakes. |
| Diet | Short text | Short marks for the menu: V for vegetarian, VG for vegan, GF for gluten free. |
| Active | Yes/No | No for a dish taken off the menu. It stays for the orders that sold it, and drops out of the menu and the costing. |

#### Suppliers

The firms the cafe buys from, one record each. The sample has 6 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Name | Short text, required | The supplier's trading name. |
| Supplies | Short text | What the cafe buys from them, in a few words. |
| Contact | Short text | The person to speak to when ordering. |
| Phone | Short text | The number to call an order through. |
| E-mail | Short text | Where to send an order in writing. |
| Delivers On | Short text | The days of the week they deliver. An order goes in the day before. |
| Payment Days | Whole number | How many days after a delivery its bill is due. |

#### Ingredients

Everything that goes into a dish or is sold as it comes, with its price, where it is kept and how much is in. The sample has 55 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Name | Short text, required, unique | The ingredient as the kitchen knows it. |
| Supplier | Points at Suppliers, required | Who it is ordered from. |
| Unit | Short text | The unit its cost, stock and recipe quantities are counted in: kg, litre or each. |
| Cost per Unit | Currency | What one unit costs now, from the latest bill. Changing it changes every dish's food cost. |
| Bought As | Short text | How it comes, such as a 1 kg bag or a tray of 30, to help when ordering. |
| Storage | Short text | Where it is kept: fridge, dry store or the counter. Used for the stock count and the stock value. |
| In Stock | Number | How much is in, in its unit. Set it from a stock count; Data Files does not take sales off it by itself. |
| On Order | Number | How much is ordered and not yet delivered. Clear it when the delivery comes in. |
| Reorder Level | Number | When stock and what is on order come to this or less, it goes on the order list. |
| Reorder Quantity | Number | The usual amount to order, in its unit. |

#### Recipe Lines

What goes into one portion of each dish: one line per ingredient. These give every dish its food cost. The sample has 160 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Menu Item | Points at Menu Items, required | The dish this line belongs to. |
| Line | Whole number | The line's place in the recipe, from 1. Every dish needs a line 1: the costing counts the price once, on that line. |
| Ingredient | Points at Ingredients, required | The ingredient used. |
| Quantity | Number | How much of it one portion uses, in the ingredient's unit: 0.018 kg of beans for a shot of espresso. |

#### Staff

One record per person on the payroll. The sample has 8 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| First Name | Short text | Their first name. |
| Last Name | Short text, required | Their family name. |
| Role | Short text | Their job, such as Head Baker or Barista. |
| Hourly Rate | Currency | What they are paid an hour before tax. The wage cost of each shift is worked out from it. |
| Phone | Short text | A number to reach them about shifts. |
| E-mail | Short text | Where the rota is sent. |
| Started | Date | The day they started. |

#### Shifts

The rota: one record per person per shift, planned ahead and kept once worked. The sample has 2,146 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Staff | Points at Staff, required | Who works the shift. |
| Start | Date and time | The day and time the shift starts. |
| End | Date and time | The day and time it ends, later than planned when a busy day kept someone on. |
| Break Minutes | Whole number | The unpaid break in the shift. |
| Paid Hours | Number | The hours paid: from start to end less the break, or nought for a day off sick. |
| Station | Short text | Where they work that shift: Barista, Kitchen, Floor, Bakery or Manager. |
| Status | Short text | Rota for a shift still to come, Worked, Holiday (paid, not worked) or Sick. |

#### Orders

One record per bill rung up on the till: a table's meal or a takeaway. The sample has 6,092 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| When | Date and time | The day and time it was rung up. |
| Service | Short text | Dine in or Takeaway. |
| Table No | Whole number | The table it was served at; empty for a takeaway. |
| Covers | Whole number | How many people sat down to it; nought for a takeaway. |
| Staff | Points at Staff, required | Who took the order. |
| Payment | Short text | Card or Cash. |
| Total | Currency | What the bill came to: the sum of its lines. |
| Voided | Yes/No | Yes for a bill cancelled before it was paid. It stays on record and is left out of the takings. |

#### Order Lines

What was on each bill: one line per dish, with how many. The sample has 16,167 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Order | Points at Orders, required | The bill this line is on. |
| Menu Item | Points at Menu Items, required | The dish or drink sold. |
| Quantity | Whole number | How many. |
| Unit Price | Currency | The price of one when it was sold, kept so a later price change leaves old bills as they were. |

#### Purchases

One record per order to a supplier, from the day it was placed to the day its bill was paid. The sample has 1,132 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Supplier | Points at Suppliers, required | Who it was ordered from. |
| Ordered On | Date | The day the order went in. |
| Due On | Date | The day the delivery is expected. |
| Received On | Date | The day it arrived; empty while it is still to come. |
| Invoice Number | Short text | The supplier's number on the bill that came with the delivery. |
| Paid On | Date | The day the bill was paid; empty while it is owed. |

#### Purchase Lines

What was ordered on each purchase: one line per ingredient. The sample has 2,903 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Purchase | Points at Purchases, required | The purchase this line is on. |
| Ingredient | Points at Ingredients, required | The ingredient ordered. |
| Quantity | Number | How much, in the ingredient's unit. |
| Unit Cost | Currency | What one unit cost on this bill. Prices rose a little over the year, so older bills are cheaper. |

#### Waste

Everything thrown away, as the kitchen notes it on the waste sheet. The sample has 647 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Ingredient | Points at Ingredients, required | What was thrown away. |
| Date | Date | The day. |
| Quantity | Number | How much, in the ingredient's unit. |
| Reason | Short text | Why: out of date, overripe, stale, spilt, dropped, a kitchen error and so on. |
| Recorded By | Points at Staff | Who wrote it down. |

#### Cash Ups

The sheet filled in at close each day the cafe opens: the till's reading, the cash counted, and the day's wages from the rota. The sample has 424 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Date | Date, required, unique | The day. One sheet per day. |
| Closed By | Points at Staff | Who counted up and locked up. |
| Weather | Short text | Sunny, Cloudy, Rain or Cold: worth noting, as it moves trade. |
| Orders | Whole number | How many bills the till rang up that day, voided ones left out. |
| Covers | Whole number | How many people sat down. |
| Card Takings | Currency | Taken by card, from the till's end of day reading. |
| Cash Takings | Currency | Taken in cash, from the same reading. |
| Cash Counted | Currency | The cash actually in the drawer, less the float. A difference from Cash Takings is the till over or short. |
| Wage Cost | Currency | What the day's shifts cost in wages, from the rota: hours worked or on holiday times each person's rate. |
| Notes | Long text | Anything worth remembering about the day: something ran out, the card machine went down, the till was out. |

### Relationships

- Menu Items.Category points at Menu Categories.ID
- Ingredients.Supplier points at Suppliers.ID
- Recipe Lines.Menu Item points at Menu Items.ID
- Recipe Lines.Ingredient points at Ingredients.ID
- Shifts.Staff points at Staff.ID
- Orders.Staff points at Staff.ID
- Order Lines.Order points at Orders.ID
- Order Lines.Menu Item points at Menu Items.ID
- Purchases.Supplier points at Suppliers.ID
- Purchase Lines.Purchase points at Purchases.ID
- Purchase Lines.Ingredient points at Ingredients.ID
- Waste.Ingredient points at Ingredients.ID
- Waste.Recorded By points at Staff.ID
- Cash Ups.Closed By points at Staff.ID

### Queries

#### Menu costing

Each dish on the menu with its price, food cost, food cost as a share of the price, and the gross profit and margin left over. Use it when setting prices or when a supplier puts a price up: most cafes aim for a food cost under about 30% of the price.

From Menu Items, Menu Categories, Recipe Lines, Ingredients. Shows Category, Menu Item, Price, Food Cost, Food Cost %, Gross Profit, Gross Margin %.
Calculated: Food Cost is `Round([Quantity] * [Cost per Unit], 2)`; Food Cost % is `Round(100 * [Quantity] * [Cost per Unit] / [Price], 2)`; Gross Profit is `Round(IIf([Line] = 1, [Price], 0) - Round([Quantity] * [Cost per Unit], 2), 2)`; Gross Margin % is `Round(IIf([Line] = 1, 100, 0) - 100 * [Quantity] * [Cost per Unit] / [Price], 2)`.
Only where `Active = Yes`.

#### Recipe card

Every line of every recipe with what it costs, for checking a recipe or showing a new cook how a dish is made up.

From Recipe Lines, Menu Items, Ingredients. Shows Menu Item, Line, Ingredient, Quantity, Unit, Cost per Unit, Line Cost.
Calculated: Line Cost is `Round([Quantity] * [Cost per Unit], 2)`.
Only where `Active = Yes`.

#### Menu

The dishes on the menu now, by section, with their diet marks and prices.

From Menu Items, Menu Categories. Shows Category, Served, Menu Item, Diet, Takeaway, Price.
Only where `Active = Yes`.

#### Best sellers, last 30 days

How many of each dish sold in the last 30 days and what they took, most sold first. Use it to plan the bake and the prep, and to spot dishes worth dropping.

From Order Lines, Orders, Menu Items, Menu Categories. Shows Menu Item, Category, Sold, Takings.
Calculated: Takings is `Round([Quantity] * [Unit Price], 2)`.
Only where `Days Ago Between 0 And 30` and `Voided = No`.

#### Sales by hour

Bills, takings and the average bill for each hour of the day over the last eight weeks: when the rush comes and when the cafe is quiet, for setting shift times.

From Orders. Shows Hour, Orders, Takings, Average Spend.
Calculated: Hour is `Mid([When], 12, 2)`.
Only where `Days Ago Between 1 And 56` and `Voided = No`.

#### Sales by day of the week

Bills, covers, takings and the average bill for each day of the week over the last eight weeks, for planning who works when.

From Orders. Shows Day, Orders, Covers, Takings, Average Spend.
Calculated: Day is `IIf(DateDiff("d", #2024-01-01#, [When]) Mod 7 = 0, "Monday", IIf(DateDiff("d", #2024-01-01#, [When]) Mod 7 = 1, "Tuesday", IIf(DateDiff("d", #2024-01-01#, [When]) Mod 7 = 2, "Wednesday", IIf(DateDiff("d", #2024-01-01#, [When]) Mod 7 = 3, "Thursday", IIf(DateDiff("d", #2024-01-01#, [When]) Mod 7 = 4, "Friday", IIf(DateDiff("d", #2024-01-01#, [When]) Mod 7 = 5, "Saturday", "Sunday"))))))`.
Only where `Days Ago Between 1 And 56` and `Voided = No`.

#### Average spend by service

For the last 30 days, how many bills were eaten in and how many taken away, what each brought in and the average bill for each.

From Orders. Shows Service, Orders, Covers, Takings, Average Spend.
Only where `Days Ago Between 0 And 30` and `Voided = No`.

#### Takings and wages by week

Each week's takings from the cash-up sheets against the wages on the rota, what was left after wages, and the average day's wages as a share of its takings. The week starts on Monday; the newest week comes first.

From Cash Ups. Shows Week Beginning, Days Open, Orders, Covers, Takings, Wages, Left After Wages, Average Day Wage %.
Calculated: Takings is `[Card Takings] + [Cash Takings]`; Left After Wages is `Round([Card Takings] + [Cash Takings] - [Wage Cost], 2)`; Average Day Wage % is `Round(100 * [Wage Cost] / ([Card Takings] + [Cash Takings]), 1)`.

#### Takings by month

Takings and wages for each month, oldest first, to see the seasons and compare with the same month last year.

From Cash Ups. Shows Month, Days Open, Orders, Takings, Wages, Left After Wages.
Calculated: Month is `Format([Date], "yyyy-mm")`; Takings is `[Card Takings] + [Cash Takings]`; Left After Wages is `Round([Card Takings] + [Cash Takings] - [Wage Cost], 2)`.

#### Takings by weather

The average day's takings and covers in the warmer months (May to September) by the weather noted on the cash-up sheet.

From Cash Ups. Shows Weather, Days, Average Takings, Average Covers.
Calculated: Average Takings is `[Card Takings] + [Cash Takings]`.
Only where `Month In ("05", "06", "07", "08", "09")`.

#### Cash over or short

The days the cash counted at close did not match the till's cash reading, and by how much, with who closed up.

From Cash Ups, Staff. Shows Date, Closed By, Cash Takings, Cash Counted, Over or Short, Notes.
Calculated: Closed By is `[First Name] & " " & [Last Name]`; Over or Short is `Round([Cash Counted] - [Cash Takings], 2)`.
Only where `Over or Short <> 0`.

#### Hours and wages by staff, last 28 days

Each person's shifts, hours and wages over the last four weeks, with holiday paid in and sick days left out. Use it for the payroll and to keep hours fair.

From Shifts, Staff. Shows Name, Role, Hourly Rate, Shifts, Hours, Wages.
Calculated: Name is `[First Name] & " " & [Last Name]`; Wages is `Round([Paid Hours] * [Hourly Rate], 2)`.
Only where `Days Ago Between 1 And 28` and `Status In ("Worked", "Holiday")`.

#### Rota for the next 7 days

Who is on, when and where, today and for the next six days, earliest first.

From Shifts, Staff. Shows Day, From, To, Name, Station, Paid Hours, Status.
Calculated: Day is `Format([Start], "yyyy-mm-dd")`; From is `Mid([Start], 12, 5)`; To is `Mid([End], 12, 5)`; Name is `[First Name] & " " & [Last Name]`.
Only where `Days Ahead Between 0 And 6`.

#### Ingredients to order

Ingredients at or below their reorder level with nothing on order, by supplier, with the usual amount to order and what it will cost. Work through it each morning before phoning the orders in.

From Ingredients, Suppliers. Shows Supplier, Delivers On, Ingredient, Unit, In Stock, On Order, Reorder Level, Order, Bought As, Estimated Cost.
Calculated: Estimated Cost is `Round([Reorder Quantity] * [Cost per Unit], 2)`.
Only where `Above Reorder Level <= 0`.

#### Deliveries due

Orders placed with suppliers and not yet delivered, the earliest due first, with how many lines and their value.

From Purchases, Suppliers, Purchase Lines. Shows Purchase, Supplier, Ordered On, Due On, Lines, Value.
Calculated: Value is `Round([Quantity] * [Unit Cost], 2)`.
Only where `Received On Is Null`.

#### Bills to pay

Delivered orders whose bills are not yet paid, with the days left until each is due; below nought means it is overdue.

From Purchases, Suppliers, Purchase Lines. Shows Purchase, Supplier, Invoice Number, Received On, Days to Pay, Amount.
Calculated: Days to Pay is `[Payment Days] - DateDiff("d", [Received On], Date())`; Amount is `Round([Quantity] * [Unit Cost], 2)`.
Only where `Received On Is Not Null` and `Paid On Is Null`.

#### Spend by supplier by month

What was spent with each supplier each month, from the delivered orders.

From Purchases, Suppliers, Purchase Lines. Shows Month, Supplier, Deliveries, Spent.
Calculated: Month is `Format([Received On], "yyyy-mm")`; Spent is `Round([Quantity] * [Unit Cost], 2)`.
Only where `Received On Is Not Null`.

#### Waste by ingredient, last 90 days

For each ingredient thrown away in the last 90 days, how often, how much and what it cost, the dearest first. Use it to order less of what goes off.

From Waste, Ingredients. Shows Ingredient, Unit, Times, Wasted, Cost.
Calculated: Cost is `Round([Quantity] * [Cost per Unit], 2)`.
Only where `Days Ago Between 0 And 90`.

#### Waste in the last 30 days

Each entry on the waste sheet in the last 30 days with its cost and who recorded it.

From Waste, Ingredients, Staff. Shows Date, Reason, Ingredient, Quantity, Unit, Cost, Recorded By.
Calculated: Cost is `Round([Quantity] * [Cost per Unit], 2)`; Recorded By is `[First Name] & " " & [Last Name]`.
Only where `Days Ago Between 0 And 30`.

#### Stock value by storage

What the stock in each store is worth at today's cost, for the month end stock count and the accounts.

From Ingredients. Shows Storage, Ingredients, Value at Cost.
Calculated: Value at Cost is `Round([In Stock] * [Cost per Unit], 2)`.

### Reports

#### Menu costing sheet

Every dish on the menu by section with its price, food cost, food cost percentage, gross profit and margin, and the average for each section. For the owner when pricing the menu.

From the query Menu costing, grouped by Category: Menu Item, Price, Food Cost, Food Cost %, Gross Profit, Gross Margin %, with average of food cost %, average of gross profit.

#### Recipe cards

A card for each dish listing its ingredients, the quantity of each and its cost, with the dish's total food cost. For the kitchen wall and for training.

From the query Recipe card, grouped by Menu Item: Line, Ingredient, Quantity, Unit, Line Cost, with total of line cost.

#### Menu

The menu by section with prices and diet marks, ready to check against the printed menu.

From the query Menu, grouped by Category: Menu Item, Diet, Price.

#### Weekly takings and wages

One line a week: days open, bills, covers, takings, wages, what was left after wages and the wage share, with the totals. For the owner's weekly look at the business.

From the query Takings and wages by week: Week Beginning, Days Open, Orders, Covers, Takings, Wages, Left After Wages, Average Day Wage %, with total of takings, total of wages, total of left after wages.

#### Monthly takings

Takings and wages month by month with the totals, for the accountant or the bank.

From the query Takings by month: Month, Days Open, Orders, Takings, Wages, Left After Wages, with total of takings, total of wages, total of left after wages.

#### Rota for the next 7 days

The coming week's shifts, a section for each day with its total hours, to pin up in the back.

From the query Rota for the next 7 days, grouped by Day: From, To, Name, Station, Paid Hours, Status, with total of paid hours.

#### Order list for suppliers

What to order today, a section for each supplier with its delivery days, the quantities and the estimated cost. For whoever phones the orders in.

From the query Ingredients to order, grouped by Supplier: Ingredient, Unit, In Stock, On Order, Reorder Level, Order, Bought As, Estimated Cost, with total of estimated cost.

#### Waste report

The last 30 days of waste grouped by reason, with the cost of each and in all. For the weekly kitchen meeting.

From the query Waste in the last 30 days, grouped by Reason: Date, Ingredient, Quantity, Unit, Cost, Recorded By, with total of cost.

#### Bills to pay

Unpaid supplier bills grouped by supplier, with their invoice numbers, the days left to pay and the total owed to each.

From the query Bills to pay, grouped by Supplier: Purchase, Invoice Number, Received On, Days to Pay, Amount, with total of amount.

### Starting empty

Start with Suppliers, then Ingredients with each one's unit, cost per unit, supplier and reorder level. Then Menu Categories, then Menu Items, then the Recipe Lines for each dish, numbering each recipe's lines from 1; Menu costing works as soon as a dish has its recipe. Add your Staff and plan the Shifts for the coming week. From then on, day to day: a Purchase and its lines when you order from a supplier, its Received On when it arrives and its Paid On when the bill is paid; the Waste sheet as things are thrown away; and a Cash Up at each close. If your till keeps its own records, the Cash Ups alone give the weekly and monthly figures; enter Orders and Order Lines only if you want best sellers and the busy hours here too. Keep In Stock up to date from a stock count, as Data Files does not take sales off it by itself.

### Making it yours

- If you charge VAT, add a VAT Rate field to Menu Items and change the costing's food cost percentage to divide by [Price] / (1 + [VAT Rate]), so it is measured against what the cafe keeps.
- Add a Packaging Cost field to Menu Items for the cup, lid or box a takeaway goes out in, and add it to the food cost in a copy of Menu costing to see what a takeaway really costs.
- Add yes/no allergen fields to Ingredients (gluten, milk, eggs, nuts and the rest) and a query that lists the dishes using any ingredient marked Yes, for the allergen folder.
- A bakery can add a Bakes table (date, dish, made, left over at close) to see how much of each cake to make on each day of the week.
- A food truck can drop Table No and Covers and add a Pitch field to Cash Ups, then group the takings by pitch to see which places pay.
- A restaurant can add a Bookings table (date, time, covers, name, phone, table) and a query for the bookings in the next 7 days.
- Add a Running Costs table (date, kind, amount) for rent, power and card fees, and a query by month beside Takings by month to see the profit.
- Add a Tips field to Cash Ups and a column summing it in Takings and wages by week, so tips are kept apart from the takings.

## Hollybrook Lettings

**Property rentals.** Buildings, units, tenants, leases, rent, repairs, safety checks and the accounts for a landlord.

Hollybrook Lettings is an invented letting business in Leeds that owns and manages twelve buildings: six houses, two blocks of flats, two houses split into flats, a parade of three shops with two flats above, and a yard of two workshops bought about a year ago and refitted. Between them they hold thirty-eight units let to tenants on residential and commercial leases.

The sample holds eighteen months of letting. Each unit has its run of leases: tenants who renewed, tenants who left, and the weeks a unit stood empty between them. Rent is charged once a month on each lease's rent day. Most tenants pay on the day by standing order, some pay late, and a few have fallen behind and pay part of what is due, so there are arrears to chase. Things break and are reported, more of them in winter, and a contractor is booked to mend each one. Homes are inspected twice a year and shops and workshops once a year; gas appliances are checked every year and the wiring every five years. Insurance, loan interest, cleaning of shared areas, letting costs and the bills for empty units are paid out of each building's account.

Money for each building goes through one Transactions table, rent in and costs out, so income and expenditure, net income and the month-by-month accounts come from one place. A rent payment also points at the month's rent charge it pays, and the charge keeps the total paid against it, which is how arrears are found.

It suits, with small changes: A private landlord with a few houses or flats; A small letting agent managing homes for owners; Holiday lets, with weekly lets in place of leases; Student houses let by the room; Storage units or lock-up garages; Allotments or moorings let by the year; A small commercial landlord with shops and workshops.

### Tables

#### Properties

Each building you own or manage, one record each. The sample has 12 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Name | Short text, required, unique | What you call the building, such as Linden Court or 22 Mill Lane. Each name is used once. |
| Kind | Short text | House, Block of flats, Converted house, Shops with flats or Workshop yard. |
| Address | Short text | The street address. |
| City | Short text | The town or city. |
| Postcode | Short text | The postcode. |
| Bought | Date | The day it was bought. |
| Purchase Price | Currency | What it cost. |
| Current Value | Currency | What it is worth now, from the last valuation. The yield query divides rent by this. |
| Valued On | Date | When it was last valued. |
| Notes | Long text | Anything worth knowing about the building: a lease on the freehold, a loan, access. |

#### Units

What is let: a house, a flat, a shop or a workshop. A house let whole is one unit; a block of flats has one unit for each flat. The sample has 38 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Property | Points at Properties, required | The building it is in. |
| Name | Short text, required | The unit's name within the building, such as Flat 3, Shop 2 or Whole house. |
| Kind | Short text | House, Flat, Shop or Workshop. |
| Bedrooms | Whole number | How many bedrooms; nought for a shop or workshop. |
| Market Rent | Currency | What the unit would let for each month now. Used for empty units and the yield. |
| Has Gas | Yes/No | Yes when the unit has gas appliances, so it needs a gas safety check every year. |
| Furnished | Yes/No | Yes when it is let with furniture. |
| Status | Short text | Let or Vacant, as it stands today. |
| Vacant Since | Date | The first day it stood empty; empty while it is let. |
| Notes | Long text | Anything about the unit: the parking space, the meter numbers, what the inventory says. |

#### Tenants

Everyone who rents from you, past and present, one record each. A tenant who renews stays one record with several leases. The sample has 63 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| First Name | Short text | Their first name, or the first name of the person who signs for a business. |
| Last Name | Short text, required | Their family name. |
| Business Name | Short text | The business that rents a shop or workshop; empty for a home. |
| Phone | Short text | Their phone number. |
| E-mail | Short text | Their e-mail address. |
| Emergency Contact | Short text | Who to call if you cannot reach them, and how they are related. |
| Emergency Phone | Short text | That person's phone number. |
| References Checked | Yes/No | Yes once their references and credit check have come back. |
| ID Checked | Yes/No | Yes once you have seen proof of who they are and their right to rent. |
| Notes | Long text | Anything worth knowing: pets, a guarantor, a forwarding address. |

#### Leases

Each agreement to let a unit to a tenant for a term. A renewal is a new lease for the same tenant and unit. The sample has 91 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Unit | Points at Units, required | The unit let. |
| Tenant | Points at Tenants, required | Who it is let to. |
| Type | Short text | Residential or Commercial. |
| Start | Date | The first day of the lease. |
| End | Date | The last day of the fixed term. |
| Monthly Rent | Currency | The rent each month. |
| Rent Day | Whole number | The day of the month rent falls due. |
| Pays By | Short text | Standing order, Bank transfer, Card or Cash. |
| Deposit | Currency | The deposit taken: five weeks' rent for a home, three months' for a shop or workshop. |
| Deposit Ref | Short text | The reference the deposit protection scheme gave for a home's deposit (invented here). Empty until it is protected. |
| Renewal | Yes/No | Yes when the lease follows an earlier one for the same tenant. |
| Status | Short text | Upcoming, Active or Ended, as it stands today. |
| Notes | Long text | Anything agreed that is not in the fields: a break clause, a rent review, a pet clause. |

#### Rent Charges

The rent due on a lease for one month: one record for each lease and month, made as each month comes round, and a month ahead. The sample has 667 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Lease | Points at Leases, required | The lease the rent is due on. |
| Period | Short text | The month it is for, as 2026-10. |
| Due Date | Date | The day it is due. |
| Amount | Currency | How much is due. |
| Amount Paid | Currency | How much has been paid against it so far. Add each payment here as you enter it in Transactions; what is left is owed. |

#### Contractors

The trades you call on for repairs and safety checks. The sample has 10 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Name | Short text, required | The firm's name. |
| Trade | Short text | What they do, such as Plumbing and heating or Electrical. |
| Contact | Short text | The person you deal with. |
| Phone | Short text | Their phone number. |
| E-mail | Short text | Where to send job details. |
| Notes | Long text | Rates, insurance, call-out hours. |

#### Maintenance Requests

Each repair asked for or found, from the day it is reported to the day it is done. The sample has 147 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Unit | Points at Units, required | The unit where the problem is. |
| Reported | Date | The day you heard of it. |
| Reported By | Short text | Tenant, Inspection or Landlord. |
| Problem | Short text, required | What is wrong, in a few words. |
| Priority | Short text | Urgent (same or next day), High (within three days), Normal or Low. |
| Status | Short text | Reported, Booked, Done or Cancelled. |
| Contractor | Points at Contractors | Who is doing the work; empty until someone is booked. |
| Booked For | Date | The day the contractor is due. |
| Completed | Date | The day the work was finished. |
| Cost | Currency | What the work cost. The same sum goes into Transactions as money out. |
| Notes | Long text | Access arrangements, parts on order, what was found. |

#### Inspections

Routine inspections of each unit and the safety checks the law asks for, each one due on a date. When one is done, the next one is entered with its due date. The sample has 248 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Unit | Points at Units, required | The unit inspected. |
| Kind | Short text | Routine inspection, Gas safety or Electrical safety. |
| Due | Date | When it must be done by. |
| Done | Date | The day it was done; empty until then. |
| Inspector | Short text | Who did it: one of your staff, or the contractor who did the check. |
| Outcome | Short text | What was found, such as Good, Action needed, Pass or Unsatisfactory. |
| Certificate Number | Short text | The number on the safety certificate (invented here). |
| Notes | Long text | Anything to know before going, or what happened. |

#### Transactions

Every sum of money in or out for a building: rent received, and every cost paid. Each building's accounts are made from these. The sample has 1,126 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Property | Points at Properties, required | The building the money is for. |
| Date | Date | The day it was paid or received. |
| Kind | Short text | Rent, Repairs, Insurance, Loan interest, Shared areas, Safety checks, Letting costs, Agent fees or Empty unit costs. |
| Description | Short text | What it was for. |
| Charge | Points at Rent Charges | For rent, the month's rent charge it pays; empty for a cost. |
| Payee or Payer | Short text | Who paid you, or who you paid. |
| Method | Short text | How it was paid. |
| Money In | Currency | What came in; nought for a cost. |
| Money Out | Currency | What went out; nought for rent. |

### Relationships

- Units.Property points at Properties.ID
- Leases.Unit points at Units.ID
- Leases.Tenant points at Tenants.ID
- Rent Charges.Lease points at Leases.ID
- Maintenance Requests.Unit points at Units.ID
- Maintenance Requests.Contractor points at Contractors.ID
- Inspections.Unit points at Units.ID
- Transactions.Property points at Properties.ID
- Transactions.Charge points at Rent Charges.ID

### Queries

#### Rent roll

Every lease running today, by building and unit, with the tenant, the term, the rent and the deposit.

From Leases, Units, Properties, Tenants. Shows Property, Unit, Tenant, Type, Start, End, Monthly Rent, Deposit.
Calculated: Tenant is `Nz([Business Name], [First Name] & " " & [Last Name])`.
Only where `Status = Active`.

#### Rent arrears by tenant

Each tenant who owes rent that is past due, with how many months are unpaid, how much they owe and how many days late the oldest month is. Former tenants who still owe are on it too.

From Rent Charges, Leases, Tenants, Units, Properties. Shows Tenant, Unit, Phone, Months Owing, Owed, Days Late.
Calculated: Tenant is `Nz([Business Name], [First Name] & " " & [Last Name])`; Unit is `[Units].[Name] & ", " & [Properties].[Name]`; Owed is `Round([Amount] - Nz([Amount Paid]), 2)`; Days Late is `DateDiff("d", [Due Date], Date())`.
Only where `Unpaid > 0` and `Overdue > 0`.

#### Unpaid rent

Every month's rent past its due day and not paid in full, by tenant, with the address to write to.

From Rent Charges, Leases, Tenants, Units, Properties. Shows Tenant, Address, Period, Due Date, Amount, Amount Paid, Owed, Days Late.
Calculated: Tenant is `Nz([Business Name], [First Name] & " " & [Last Name])`; Address is `IIf([Properties].[Name] = [Address], IIf([Units].[Name] = "Whole house", [Address], [Units].[Name] & ", " & [Address]), [Units].[Name] & ", " & [Properties].[Name] & ", " & [Address]) & ", " & [Postcode]`; Owed is `Round([Amount] - Nz([Amount Paid]), 2)`; Days Late is `DateDiff("d", [Due Date], Date())`.
Only where `Owed > 0` and `Days Late > 0`.

#### Rent due in the next 14 days

Rent falling due in the next two weeks and not yet paid, with how each tenant pays, so you know what to expect in the bank.

From Rent Charges, Leases, Tenants, Units, Properties. Shows Due Date, Days Until, Tenant, Unit, Pays By, Amount.
Calculated: Days Until is `DateDiff("d", Date(), [Due Date])`; Tenant is `Nz([Business Name], [First Name] & " " & [Last Name])`; Unit is `[Units].[Name] & ", " & [Properties].[Name]`.
Only where `Days Until Between 0 And 14` and `Owed > 0`.

#### Leases ending in the next 60 days

Leases whose term ends within two months, with the rent now and what the unit would let for, so you can offer a renewal or plan to relet.

From Leases, Units, Properties, Tenants. Shows Property, Unit, Tenant, Phone, End, Days Left, Monthly Rent, Market Rent.
Calculated: Tenant is `Nz([Business Name], [First Name] & " " & [Last Name])`; Days Left is `DateDiff("d", Date(), [End])`.
Only where `Days Left Between 0 And 60` and `Status = Active`.

#### Vacant units

Units standing empty, how long for, and the rent lost at the market rent since they emptied.

From Units, Properties. Shows Property, Unit, Kind, Bedrooms, Market Rent, Vacant Since, Days Empty, Rent Lost.
Calculated: Days Empty is `DateDiff("d", [Vacant Since], Date())`; Rent Lost is `Round([Market Rent] * 12 / 365 * DateDiff("d", [Vacant Since], Date()), 2)`.
Only where `Status = Vacant`.

#### Open repairs

Repairs not yet done or cancelled, the most urgent first and then the oldest, with how many days each has been open.

From Maintenance Requests, Units, Properties. Shows Job, Priority, Unit, Problem, Reported, Days Open, Status.
Calculated: Unit is `[Units].[Name] & ", " & [Properties].[Name]`; Days Open is `DateDiff("d", [Reported], Date())`.
Only where `Status Not In ("Done", "Cancelled")`.

#### Repairs booked

Jobs booked with a contractor and not yet done, by contractor and day, with the address and the problem.

From Maintenance Requests, Units, Properties, Contractors. Shows Contractor, Phone, Job, Booked For, Priority, Address, Problem.
Calculated: Address is `IIf([Properties].[Name] = [Address], IIf([Units].[Name] = "Whole house", [Address], [Units].[Name] & ", " & [Address]), [Units].[Name] & ", " & [Properties].[Name] & ", " & [Address]) & ", " & [Postcode]`.
Only where `Status = Booked`.

#### Repair costs by contractor

What each contractor has been paid for finished jobs in the last twelve months, and how many jobs.

From Maintenance Requests, Contractors. Shows Contractor, Trade, Jobs, Spent.
Only where `Status = Done` and `Days Ago Between 0 And 365`.

#### Inspections due

Inspections and safety checks due in the next 30 days or already overdue (a Days Left below nought is overdue).

From Inspections, Units, Properties. Shows Due, Days Left, Kind, Unit, Notes.
Calculated: Days Left is `DateDiff("d", Date(), [Due])`; Unit is `[Units].[Name] & ", " & [Properties].[Name]`.
Only where `Days Left <= 30` and `Done Is Null`.

#### Safety certificates due

Gas and electrical safety checks due in the next 60 days or overdue, so the engineer can be booked in time.

From Inspections, Units, Properties. Shows Due, Days Left, Kind, Unit.
Calculated: Days Left is `DateDiff("d", Date(), [Due])`; Unit is `[Units].[Name] & ", " & [Properties].[Name]`.
Only where `Days Left <= 60` and `Kind In ("Gas safety", "Electrical safety")` and `Done Is Null`.

#### Deposits held

The deposit taken on every lease running today, with its protection reference.

From Leases, Tenants, Units, Properties. Shows Property, Unit, Tenant, Type, Deposit, Deposit Ref.
Calculated: Tenant is `Nz([Business Name], [First Name] & " " & [Last Name])`.
Only where `Status = Active` and `Deposit > 0`.

#### Deposits to protect

Home deposits not yet protected, with how many days are left of the 30 allowed from the start of the lease.

From Leases, Tenants, Units, Properties. Shows Tenant, Unit, Start, Status, Deposit, Days to Protect.
Calculated: Tenant is `Nz([Business Name], [First Name] & " " & [Last Name])`; Unit is `[Units].[Name] & ", " & [Properties].[Name]`; Days to Protect is `30 - DateDiff("d", [Start], Date())`.
Only where `Type = Residential` and `Deposit Ref Is Null` and `Status In ("Active", "Upcoming")`.

#### Current tenants

Everyone living or trading in your units today, with their phone, e-mail and emergency contact.

From Leases, Tenants, Units, Properties. Shows Property, Unit, Tenant, Phone, E-mail, Emergency Contact, Emergency Phone.
Calculated: Tenant is `Nz([Business Name], [First Name] & " " & [Last Name])`.
Only where `Status = Active`.

#### Income and expenditure by property

For each building over the last twelve months: rent received, costs paid and the net income.

From Transactions, Properties. Shows Property, Money In, Money Out, Net.
Calculated: Net is `Round(Nz([Money In]) - Nz([Money Out]), 2)`.
Only where `Days Ago Between 0 And 365`.

#### Costs by kind

What was spent in the last twelve months on each kind of cost, the biggest first.

From Transactions. Shows Kind, Payments, Spent.
Only where `Money Out > 0` and `Days Ago Between 0 And 365`.

#### Money in and out by month

Money in, money out and the net for each month of the last year.

From Transactions. Shows Month, Money In, Money Out, Net.
Calculated: Month is `Format([Date], "yyyy-mm")`; Net is `Round(Nz([Money In]) - Nz([Money Out]), 2)`.
Only where `Months Ago Between 0 And 12`.

#### Gross yield by property

Each building's rent for a year at market rents as a percentage of its current value, the highest first.

From Properties, Units. Shows Property, Kind, Units, Value, Annual Rent, Gross Yield %.
Calculated: Annual Rent is `Round([Market Rent] * 12, 2)`; Gross Yield % is `Round([Market Rent] * 1200 / [Current Value], 2)`.

### Reports

#### Rent roll

The rent roll: each building's units, tenants, terms, rents and deposits, with totals for each building and for all of them. The page a lender or buyer asks for.

From the query Rent roll, grouped by Property: Unit, Tenant, Start, End, Monthly Rent, Deposit, with total of monthly rent, total of deposit.

#### Rent arrears

Each tenant in arrears on their own section, with every unpaid month and the total owed: the list to write the reminder letters from.

From the query Unpaid rent, grouped by Tenant: Address, Period, Due Date, Days Late, Amount, Amount Paid, Owed, with total of owed.

#### Jobs booked

A page for each contractor listing the jobs booked with them, the day, the address and the problem, to send or hand over.

From the query Repairs booked, grouped by Contractor: Job, Booked For, Priority, Address, Problem.

#### Income and expenditure, last 12 months

Money in, money out and the net for each building over the last twelve months, with the totals: the figures for your accounts.

From the query Income and expenditure by property: Property, Money In, Money Out, Net, with total of money in, total of money out, total of net.

#### Leases ending in the next 60 days

Leases ending in the next 60 days with the tenant's phone, the rent now and the market rent.

From the query Leases ending in the next 60 days: Property, Unit, Tenant, Phone, End, Days Left, Monthly Rent, Market Rent, with total of monthly rent.

#### Vacant units and rent lost

The units standing empty, how long for and the rent lost on each.

From the query Vacant units: Property, Unit, Kind, Market Rent, Vacant Since, Days Empty, Rent Lost, with total of rent lost, total of market rent.

#### Inspections and safety checks due

Inspections and safety checks due in the next 30 days or overdue, grouped by kind, to plan the visits.

From the query Inspections due, grouped by Kind: Due, Days Left, Unit, Notes.

#### Monthly accounts

One line a month for the last year with money in, money out and the net, and the totals.

From the query Money in and out by month: Month, Money In, Money Out, Net, with total of money in, total of money out, total of net.

### Starting empty

Start with Properties, then the Units in each one, then Contractors. Enter each Tenant and their Lease, and for every unit the next inspection and safety checks due. Each month add a Rent Charge for every running lease; when rent comes in, enter it in Transactions pointing at the charge, and add it to the charge's Amount Paid. Enter repairs in Maintenance Requests as they are reported, and every bill you pay in Transactions with its building.

### Making it yours

- Manage for owners: add an Owners table with their management fee, point Properties at it, and add a query that takes your fee off each owner's rent to show what you pay them.
- Let holiday homes: use a weekly Rent Charge in place of a monthly one, rename Leases to Bookings with arrival and departure dates, and add a cleaning cost to each booking.
- Let rooms in shared houses: make each room a unit, and add a field to Units for the shared bathroom and kitchen it uses.
- Let storage units or garages: drop Bedrooms and Has Gas from Units, add a Size field, and keep the rest as it is.
- Keep the energy rating: add an Energy Rating and an Energy Certificate Expires field to Units, and a query for certificates running out.
- Record rent reviews on commercial leases: add a Next Review date to Leases and a query for reviews due in the next 90 days.
- Recharge damage to tenants: add a Recharge field to Maintenance Requests, and take it from the deposit when the lease ends.

## Lantern Learning Centre

**School or training centre.** Students, families, classes, registers, test results and termly fees for a tuition or training centre.

For a centre that teaches groups in terms and bills by the term: a tuition business, a language or music school, a dance studio, a sports academy. It keeps who comes, who pays, which class each student is in, who was there at each lesson, how they did in their tests and what is still owed.

Lantern Learning Centre is an invented tuition centre in Leeds with about a hundred and twenty students. After school and on Saturday mornings it teaches children from Year 3 to Year 11 (ages seven to sixteen) in groups of eight to ten: maths, English and science for each stage up to the GCSE exams, and coding. On weekday evenings adults come for Spanish, French and Mandarin, and for bookkeeping in the computer room. In the school holidays it runs daytime workshops: a maths booster, GCSE revision days, a coding club, a reading and writing club and science lab days. Ten teachers are paid by the hour.

The records go back to the start of a term fourteen to eighteen months before the day the file is made: three terms a year with the holidays between, run to the end of the current one, with the next term or holiday planned when it starts within six weeks. Each class has its lessons on one weekday at one time, a register for every lesson that has happened, and tests twice a term for the children. Families get one invoice a term, three weeks before it starts, with ten percent off for a second or third child; most pay before the first lesson, some pay in three instalments, some pay late and a few still owe from earlier terms.

Children's records hold only what a centre needs: a name, the date of birth, the school year, a short note such as an allergy or an inhaler, and whether photographs may be taken. Addresses and phone numbers belong to the family, not the child. All of it is invented.

It suits, with small changes: A tutoring business with group or one-to-one lessons; A language school with evening classes; A music or dance school that teaches in terms; A driving school, with a class for each instructor's block of lessons; A sports academy or swimming club with weekly sessions; A small private or weekend school; An adult education or evening class programme.

### Tables

#### Terms

The blocks of weeks the centre teaches in: three terms a year and the school holidays between them, when holiday workshops run. Every class belongs to one, and fees are invoiced by it. Enter these first. The sample has 7 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Name | Short text, required, unique | The name families know the block by, such as Autumn 2026 or Easter holiday 2026. |
| Kind | Short text | Term for a teaching term, Holiday for the holiday workshops between terms. |
| Starts | Date, required | The first day of the block. Invoices go out three weeks before it. |
| Ends | Date, required | The last day of the block. A class whose term has not ended is current or coming, and the class lists and free places look only at those. |

#### Families

One record for each household that pays: the parent or carer who is the centre's contact, and where to send letters. Brothers and sisters share one family. An adult learner is their own family, with Relationship set to Self. The sample has 186 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Name | Short text, required | The name the account goes by, such as Okafor family, or the adult learner's own name. |
| Contact | Short text | The person the centre talks to and sends invoices to. |
| Relationship | Short text | What the contact is to the students: Parent, Grandparent, Carer, or Self for an adult learner. |
| Phone | Short text | The contact's phone number, used when a child is unwell or a class is cancelled. |
| E-mail | Short text | Where invoices and reports are sent. |
| Second Contact | Short text | Another adult who may collect the children or be called, when there is one. |
| Second Phone | Short text | The second contact's phone number. |
| Address | Short text | House number and street, for letters. |
| City | Short text | The town or city. |
| Postcode | Short text | The postcode. |
| Pays By | Short text | How the family usually pays: Bank transfer, Card or Cash. |
| Notes | Long text | Anything the office should know, such as a bursary place or who collects the children. |

#### Students

One record per learner, child or adult. Keep a child's record to what the centre needs; contact details are on the family. The sample has 213 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Family | Points at Families, required | The family who pays and is contacted for this student. |
| First Name | Short text | The student's first name, as used in class. |
| Last Name | Short text, required | The family name; lists are sorted by it. |
| Date of Birth | Date | For children, to place them in the right school year and for the birthdays list. Left empty for adults. |
| School Year | Short text | The school year the child is in this year, Year 3 to Year 11, or Adult. Change it each September as children move up. |
| Medical Note | Short text | A short note for the teacher, such as a nut allergy or an inhaler kept in the bag. Empty for most students. |
| Photo Consent | Yes/No | Yes when the family has agreed that photographs of the student may be taken at the centre. |
| Joined | Date | The day the student first enrolled. |
| Active | Yes/No | Yes while the student is in a class now or was in one in the last two months. Set it to No when they leave; their records stay. |

#### Teachers

The people who teach, paid by the hour for the lessons they take. The sample has 10 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| First Name | Short text | The teacher's first name. |
| Last Name | Short text, required | The teacher's family name. |
| Subjects | Short text | What they teach, such as Maths, Science. Used to find cover when a teacher is away. |
| Hourly Rate | Currency | What the centre pays them for an hour of teaching. |
| Phone | Short text | The teacher's phone number. |
| E-mail | Short text | The teacher's e-mail address at the centre. |
| Started | Date | The day they started teaching at the centre. |
| Background Check | Date | The date of their latest background check for working with children. The centre renews them every three years. |
| Active | Yes/No | Yes while they teach at the centre. |

#### Courses

What the centre teaches: one record for each subject at each level, with its length and price. A course runs as one or more classes each term. The sample has 25 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Code | Short text, unique | A short code, such as MA56 for maths for Years 5 and 6. Class codes start with it. |
| Name | Short text, required | The course name on the timetable and invoices. |
| Subject | Short text | Maths, English, Science, Coding, Spanish, French, Mandarin or Bookkeeping. Teachers are matched to classes by it. |
| Level | Short text | Who it is for: the school years, GCSE, Beginner or Improver. |
| For | Short text | Children or Adults. |
| Holiday Course | Yes/No | Yes for the workshops that run in the school holidays. |
| Minutes | Whole number | How long each lesson lasts. |
| Fee per Lesson | Currency | The price of one lesson. A term's fee is this times the lessons billed. |
| Capacity | Whole number | The most students a class of this course takes. |

#### Classes

One group of students taking a course in one term, on one weekday at one time, in one room, with one teacher. The sample has 120 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Code | Short text, required | The course code, the weekday and the start time, such as MA56 Tue 16:00. |
| Course | Points at Courses, required | The course this class teaches. |
| Teacher | Points at Teachers | The class's own teacher. A lesson taken by someone else records them on the lesson. |
| Term | Points at Terms, required | The term or holiday the class runs in. |
| Weekday | Short text | The day of the week it meets. |
| Start Time | Short text | When the lesson starts, as 16:00. |
| Room | Short text | Where it meets: Room 1, Room 2, Room 3 or the Computer Room. |
| Capacity | Whole number | How many students it takes, set by the course and the room. |
| Lessons Planned | Whole number | How many lessons it has this term, leaving out the days the centre is closed. |

#### Enrolments

One record for each student in each class: the place booked and what is charged for it. The sample has 701 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Student | Points at Students, required | Who is enrolled. |
| Class | Points at Classes, required | The class they are enrolled in. |
| Enrolled On | Date | The day the place was booked. |
| Status | Short text | Enrolled for a class still running or about to start, Completed when the term has ended, Withdrawn when the student left part way through. |
| Lessons Billed | Whole number | How many lessons are charged for: all of them, or those left for a student who joined after the start. |
| Discount | Number | A share off the fee, as 0.1 for ten percent: brothers and sisters after the first get 0.1, a bursary place 0.5. |
| Left On | Date | The day a withdrawn student left. Fees are not refunded part way through a term. |

#### Lessons

Every lesson of every class: those taught, those planned for the rest of the term and the next, and the few cancelled. The sample has 1,483 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Class | Points at Classes, required | The class the lesson is for. |
| Teacher | Points at Teachers | Who took the lesson, which is the class's teacher unless someone covered. Teacher pay is worked out from it. |
| Lesson Date | Date, required | The day of the lesson. |
| Minutes | Whole number | How long it lasted, which is what the teacher is paid for. |
| Topic | Short text | What the lesson covered, or is planned to cover. |
| Status | Short text | Taught, Planned or Cancelled. Only taught lessons are paid. |

#### Attendance

The register: one mark for each student at each lesson taught. The sample has 7,485 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Lesson | Points at Lessons, required | The lesson the mark is for. |
| Student | Points at Students, required | The student marked. |
| Mark | Short text | Present, Late or Absent. |
| Reason | Short text | Why the student was absent or late, when the family said. |

#### Assessments

The tests and marked pieces of work set to a class. Children's classes have a test at half term and one at the end of term; adult courses have one at the end. The sample has 143 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Class | Points at Classes, required | The class that sat it. |
| Assessment Date | Date | The day it was sat or handed in. |
| Title | Short text | What it was, such as Half-term check or End of term test. |
| Out Of | Whole number | The highest mark possible. |

#### Results

Each student's mark in each assessment they sat. A student away that day has no result. The sample has 838 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Assessment | Points at Assessments, required | The assessment the mark is for. |
| Student | Points at Students, required | Who sat it. |
| Score | Whole number | The mark, out of the assessment's Out Of. |
| Comment | Short text | The teacher's short comment, when there is one; it goes on the progress report. |

#### Fee Invoices

One invoice for each family for each term or holiday they book, covering every class their students take in it. The sample has 464 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Family | Points at Families, required | Who the invoice is for. |
| Term | Points at Terms, required | The term or holiday it is for. |
| Invoice Date | Date | The day it was sent, three weeks before the term starts or when a late place was booked. |
| Due Date | Date | When it should be paid: by the first day of the term. |
| Amount | Currency | The total of the family's fees for the term, after discounts. |
| Paid So Far | Currency | What has been paid against it. Keep it equal to the total of its payments; what is left is outstanding. |

#### Payments

Each payment received against an invoice. A family paying in instalments has several. The sample has 597 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Invoice | Points at Fee Invoices, required | The invoice the payment is for. |
| Paid On | Date | The day the money arrived. |
| Amount | Currency | How much was paid. |
| Method | Short text | Bank transfer, Card or Cash. |

### Relationships

- Students.Family points at Families.ID
- Classes.Course points at Courses.ID
- Classes.Teacher points at Teachers.ID
- Classes.Term points at Terms.ID
- Enrolments.Student points at Students.ID
- Enrolments.Class points at Classes.ID
- Lessons.Class points at Classes.ID
- Lessons.Teacher points at Teachers.ID
- Attendance.Lesson points at Lessons.ID
- Attendance.Student points at Students.ID
- Assessments.Class points at Classes.ID
- Results.Assessment points at Assessments.ID
- Results.Student points at Students.ID
- Fee Invoices.Family points at Families.ID
- Fee Invoices.Term points at Terms.ID
- Payments.Invoice points at Fee Invoices.ID

### Queries

#### Class lists

Every student in every class that is running now or about to start, with the class's day, time, room and teacher, any medical note, photo consent and the family's phone number. Print it as the register before a lesson.

From Enrolments, Students, Classes, Courses, Terms, Teachers, Families. Shows Class, Course, Weekday, Start Time, Room, Teacher, Student, School Year, Medical Note, Photo Consent, Contact, Phone.
Calculated: Class is `[Classes].[Code] & ", " & [Terms].[Name]`; Teacher is `[Teachers].[First Name] & " " & [Teachers].[Last Name]`; Student is `[Students].[First Name] & " " & [Students].[Last Name]`.
Only where `Status = Enrolled` and `Days to End >= 0`.

#### Timetable, next 7 days

Every lesson from today to a week ahead, in date and time order, with the room, the teacher and the planned topic. Use it to plan rooms and cover.

From Lessons, Classes, Courses, Teachers. Shows Lesson Date, Weekday, Start Time, Course, Class, Room, Teacher, Minutes, Topic.
Calculated: Teacher is `[Teachers].[First Name] & " " & [Teachers].[Last Name]`.
Only where `Days Ahead Between 0 And 7` and `Status <> "Cancelled"`.

#### Attendance below 85 percent

Students who have missed more than fifteen in a hundred of their lessons in the last 90 days, with at least four lessons in that time, and the family to call.

From Attendance, Students, Lessons, Families. Shows Student, School Year, Contact, Phone, Lessons, Absences, Attendance %.
Calculated: Student is `[Students].[First Name] & " " & [Students].[Last Name]`; Absences is `IIf([Mark] = "Absent", 1, 0)`; Attendance % is `IIf([Mark] = "Absent", 0, 100)`.
Only where `ID >= 4` and `Attendance % < 85` and `Days Ago Between 1 And 90`.

#### Fees outstanding by family

Each family that still owes something, with how many invoices are open, the oldest due date and the total left to pay, largest first.

From Fee Invoices, Families. Shows Family, Contact, Phone, Invoices, Oldest Due, Outstanding.
Calculated: Outstanding is `Round([Amount] - Nz([Paid So Far]), 2)`.
Only where `Balance > 0`.

#### Overdue fees

Each invoice past its due date that is not fully paid, with the family's address, the term, how many days late it is and what is left to pay. The fees letters are written from it.

From Fee Invoices, Families, Terms. Shows Family, Contact, Address, City, Postcode, Invoice, Term, Due Date, Amount, Days Overdue, Outstanding.
Calculated: Days Overdue is `DateDiff("d", [Due Date], Date())`; Outstanding is `Round([Amount] - Nz([Paid So Far]), 2)`.
Only where `Days Overdue > 0` and `Outstanding > 0`.

#### Classes with free places

Classes running now or next with room for more students: places, how many are enrolled and how full the class is. Use it when a family asks for a place.

From Classes, Courses, Terms, Teachers, Enrolments. Shows Class, Term, Course, Weekday, Start Time, Teacher, Places, Enrolled, Percent Full.
Calculated: Teacher is `[Teachers].[First Name] & " " & [Teachers].[Last Name]`; Percent Full is `100 / [Classes].[Capacity]`.
Only where `Percent Full < 100` and `Status = Enrolled` and `Days to End >= 0`.

#### Results by course

For each course, how many test results there were in the last twelve months and the average, highest and lowest marks as percentages.

From Results, Assessments, Classes, Courses. Shows Course, Subject, Results, Average %, Highest %, Lowest %.
Calculated: Average % is `Round([Score] * 100 / [Out Of], 0)`; Highest % is `Round([Score] * 100 / [Out Of], 0)`; Lowest % is `Round([Score] * 100 / [Out Of], 0)`.
Only where `Days Ago Between 0 And 365`.

#### Progress by student

Every test result of the last twelve months, student by student, with the mark, the percentage and the teacher's comment. The progress report prints it.

From Results, Students, Assessments, Classes, Courses. Shows Student, Course, Assessment, Assessment Date, Score, Out Of, Percent, Comment.
Calculated: Student is `[Students].[First Name] & " " & [Students].[Last Name] & " (" & [School Year] & ")"`; Percent is `Round([Score] * 100 / [Out Of], 0)`.
Only where `Days Ago Between 0 And 365`.

#### Enrolments by course and term

How many places were booked on each course in each term and holiday, and the fees charged for them.

From Enrolments, Classes, Courses, Terms. Shows Starts, Term, Course, Enrolments, Fees.
Calculated: Fees is `Round([Fee per Lesson] * [Lessons Billed] * (1 - Nz([Discount])), 2)`.

#### Teacher hours last month

For each teacher, the lessons they taught last month, the hours and what they are owed at their hourly rate. Run it at the start of each month for the pay run.

From Lessons, Teachers. Shows Month, Teacher, Hourly Rate, Lessons, Hours, Pay.
Calculated: Month is `Format([Lesson Date], "yyyy-mm")`; Teacher is `[Teachers].[First Name] & " " & [Teachers].[Last Name]`; Hours is `[Minutes] / 60`; Pay is `Round([Minutes] / 60 * [Hourly Rate], 2)`.
Only where `Status = Taught` and `Months Ago = 1`.

#### Fees received by month

Money received each month from fees, and how many payments made it up.

From Payments. Shows Month, Payments, Received.
Calculated: Month is `Format([Paid On], "yyyy-mm")`.

#### Birthdays coming up

Active students with a birthday in about the next month, with the age they turn. Adults have no date of birth here, so only children appear.

From Students. Shows Student, School Year, Birthday, Turns.
Calculated: Student is `[First Name] & " " & [Last Name]`; Birthday is `Format([Date of Birth], "dd") & "/" & Format([Date of Birth], "mm")`; Turns is `Year(Date()) - Year([Date of Birth]) + IIf(Month([Date of Birth]) * 31 + Day([Date of Birth]) < Month(Date()) * 31 + Day(Date()), 1, 0)`.
Only where `Days Away Between 0 And 30` and `Active = Yes`.

#### Background checks due

Teachers whose background check is 33 months old or more, so it can be renewed before three years are up.

From Teachers. Shows Teacher, Subjects, Background Check, Months Old.
Calculated: Teacher is `[First Name] & " " & [Last Name]`; Months Old is `DateDiff("m", [Background Check], Date())`.
Only where `Months Old >= 33` and `Active = Yes`.

### Reports

#### Class register

A page for each class running now or next: the students with their school year, medical note, photo consent and the family's contact and phone. Teachers take it into the lesson.

From the query Class lists, grouped by Class: Student, School Year, Medical Note, Photo Consent, Contact, Phone, with count of student.

#### Timetable for the next 7 days

The lessons of the coming week, day by day and in time order, with course, class, room and teacher. Pin it up in the office.

From the query Timetable, next 7 days, grouped by Lesson Date: Start Time, Course, Class, Room, Teacher, Minutes, with count of class.

#### Fees overdue

Overdue invoices family by family, with the address to write to and the total owed. Use it to send reminder letters.

From the query Overdue fees, grouped by Family: Contact, Address, City, Postcode, Invoice, Term, Due Date, Days Overdue, Outstanding, with total of outstanding.

#### Progress report

Each student's test results over the last year with the teacher's comments and their average percentage, to send home at the end of term.

From the query Progress by student, grouped by Student: Course, Assessment, Assessment Date, Score, Out Of, Percent, Comment, with average of percent.

#### Teacher pay

Hours and pay for each teacher for last month, with the totals, for the pay run.

From the query Teacher hours last month: Month, Teacher, Hourly Rate, Lessons, Hours, Pay, with total of hours, total of pay.

#### Attendance concerns

The students whose attendance has fallen below 85 percent, with the family to call.

From the query Attendance below 85 percent: Student, School Year, Contact, Phone, Lessons, Absences, Attendance %.

#### Enrolments by course

For each course, the places booked and the fees charged term by term, with totals.

From the query Enrolments by course and term, grouped by Course: Term, Starts, Enrolments, Fees, with total of enrolments, total of fees.

#### Fees received

Fee income month by month, with the total.

From the query Fees received by month: Month, Payments, Received, with total of payments, total of received.

### Starting empty

Start with Terms (this year's terms and holidays), then Teachers and Courses (what you teach, how long and at what price). Then set up the Classes for the first term, each with its course, teacher, day, time and room. Enter Families and then their Students as they book, and an Enrolment for each place. Add the Lessons of each class (one for each week of the term), take the register in Attendance as the lessons happen, and record Assessments and Results when tests are marked. Raise a Fee Invoice for each family each term and enter Payments as they come in, adding each to the invoice's Paid So Far.

### Making it yours

- For one-to-one tutoring, set each course's Capacity to 1 and make a class for each student's regular slot; the timetable and teacher pay work as they are.
- For a music or dance school, use the subjects for instruments or dance styles and the levels for grades; add a Grade field to Results for exam grades.
- If you charge by the month instead of by the term, make each month a Term record and raise the invoices monthly.
- Add a Waiting status to Enrolments for full classes, and a query listing waiting students by class, oldest first.
- For a driving school, make each learner's block of lessons a class with Capacity 1 and the instructor as the teacher; add a Test Date field to Students.
- To keep a fuller register, add a Minutes Late field to Attendance, or a Homework Done yes/no field.
- If teachers are paid a fixed amount a lesson instead of by the hour, add a Rate per Lesson field to Teachers and change the Pay column of Teacher hours last month to use it.
- Add a Rooms table with each room's size if you have many rooms, and point Classes at it instead of typing the room.

## Rowanmere Community Trust

**Charity or club.** Members, supporters, gifts, tax claims, events, volunteers and grants for a small community charity or club.

For a small charity or club run by a committee and a part-time treasurer: one database for the people who belong, give, come to events and help out, and for the money that comes in.

Rowanmere Community Trust is an invented community charity in Glasgow. It has about two hundred members who pay a yearly fee, regular supporters who give by standing order, appeals for things like a minibus or a new garden path, a hall where it holds a quiz, a ceilidh or a coffee morning most fortnights, a few dozen volunteers, and grants from a handful of invented trusts and the local council.

The records cover the last eighteen months and run into the coming weeks: memberships due for renewal, events with bookings already taken, the volunteer rota and grant reports still to write. Gifts from people who have made a tax declaration are gathered into a claim each quarter, as a UK charity claims Gift Aid; the current quarter's gifts are waiting for the next claim. Every payment in is also written once in the Income table, so the treasurer can see income by source and by month.

It suits, with small changes: A sports or social club with yearly subscriptions; A local history, gardening or arts society; A parent teacher association; A place of worship with regular givers and volunteers; A community centre or village hall committee; A small animal rescue or food bank run by volunteers.

### Tables

#### People

One record per person the trust knows: members, donors, event guests and volunteers are all here, once each. What someone is to the trust comes from their other records (a membership, a gift, a shift). The sample has 865 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| First Name | Short text | The given name, used on letters and lists. |
| Last Name | Short text, required | The family name; lists are sorted by it. |
| Address | Short text | House number and street, for letters and for a tax claim, which needs the house and postcode of each giver. |
| City | Short text | The town or city. |
| Postcode | Short text | The postcode or ZIP code. |
| Country | Short text | The country, United Kingdom unless you enter another. |
| Phone | Short text | A phone number, used for the volunteer rota and event changes. |
| E-mail | Short text | An e-mail address, left empty for people who have none. |
| Joined | Date | The day the person first came onto the records, by joining, giving or booking. |
| Contact by Post | Yes/No | Yes when the person has agreed to letters and the newsletter by post. The mailing list uses it. |
| Contact by E-mail | Yes/No | Yes when the person has agreed to e-mail from the trust. |
| Contact by Phone | Yes/No | Yes when the person has agreed to be phoned. |
| Tax Declaration | Yes/No | Yes when the person has declared that they pay enough tax for the trust to claim tax back on their gifts (in the UK, a Gift Aid declaration). Many countries have a scheme like it; where yours has none, ignore it. |
| Declaration Date | Date | The day the declaration was made. Gifts from that day on are claimable. |
| Notes | Long text | Anything worth remembering, such as a gift made in memory of someone or a request for large print. |

#### Membership Types

The kinds of membership on offer and what each costs. Enter these first. The sample has 5 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Name | Short text, required, unique | The name of the membership, such as Adult or Family. |
| Fee | Currency | What the membership costs now. A renewal is charged at this fee. |
| Months | Whole number | How long it lasts, usually 12. Nought means it never ends, as for a life membership. |
| Description | Short text | Who the membership is for. |

#### Memberships

One record for each year of membership paid. A renewal is a new record starting the day after the last one ends, so a member's history stays. The sample has 497 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Person | Points at People, required | The member. |
| Membership Type | Points at Membership Types, required | The kind of membership. |
| Starts | Date | The first day of the membership. |
| Ends | Date | The last day of the membership; empty for a life membership. |
| Amount | Currency | What was paid for it. |
| Paid On | Date | The day the fee was paid. |
| Method | Short text | How it was paid: standing order, card, cash or bank transfer. |
| Renewed | Yes/No | Yes once a later membership has been taken out, so this one no longer shows as due or lapsed. |

#### Campaigns

Appeals for a stated purpose, each with a target, and the General Fund for gifts given to no appeal in particular. Every gift belongs to one. The sample has 6 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Name | Short text, required, unique | The name of the appeal or fund. |
| Opens | Date | The day the appeal started. |
| Closes | Date | The day it ended or ends; empty for a fund that stays open. |
| Target | Currency | How much the appeal set out to raise; empty for the General Fund. |
| Description | Short text | What the money is for. |

#### Tax Claims

Each claim made to the tax office for tax back on gifts, made each quarter. The gifts in a claim point at it. The sample has 6 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Reference | Short text, unique | The trust's own reference for the claim, such as TC-004. |
| Gifts Up To | Date | The last gift date the claim covers. |
| Submitted On | Date | The day the claim was sent. |
| Amount Claimed | Currency | The total claimed: a quarter of the gifts in it, the rate the UK scheme pays. |
| Received On | Date | The day the money arrived; empty while it is awaited. |
| Amount Received | Currency | What was paid. |

#### Donations

One record per gift, whether one-off or a monthly standing order. The sample has 1,703 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Person | Points at People, required | Who gave it. |
| Campaign | Points at Campaigns, required | The appeal or fund it was given to. |
| Gift Date | Date | The day it was received. |
| Amount | Currency | How much was given. |
| Method | Short text | How: standing order, card, online, cash, cheque or bank transfer. |
| Regular | Yes/No | Yes for a gift that comes every month. |
| Claimable | Yes/No | Yes when the giver had a tax declaration in place on the day of the gift. |
| Tax Claim | Points at Tax Claims | The claim the gift went into; empty until it is claimed. |

#### Events

Everything the trust holds: quiz nights, ceilidhs, coffee mornings, talks, the summer fair and the AGM, past and planned. The sample has 68 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Name | Short text, required | What the event is. |
| Event Date | Date | The day it is held. |
| Start Time | Short text | When it starts, as 19:30. |
| Venue | Short text | Where it is held. |
| Capacity | Whole number | How many people the venue takes. |
| Ticket Price | Currency | The price of one ticket; nought for a free event. |
| Costs | Currency | What the event cost to put on: hall hire, a band, food. |
| Status | Short text | Planned, Held or Cancelled. |

#### Event Bookings

One record per booking: a person and the tickets they took for one event. The sample has 1,564 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Event | Points at Events, required | The event booked. |
| Person | Points at People, required | Who booked. |
| Booked On | Date | The day the booking was made and paid. |
| Tickets | Whole number | How many tickets. |
| Amount Paid | Currency | What was paid for them; nought once a booking is cancelled and refunded. |
| Cancelled | Yes/No | Yes when the booking was cancelled, or the event was. |
| Attended | Yes/No | Yes when the party came on the day, ticked at the door. |

#### Volunteer Roles

The jobs volunteers do, with a value per hour so their time can be counted in money for funders and annual reports. The sample has 6 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Name | Short text, required, unique | The role. |
| Usual Day | Short text | The day of the week the role usually runs. |
| Hour Value | Currency | What an hour of the role is worth if it were paid for, at a rate you choose. |
| Description | Short text | What the role involves. |

#### Volunteer Shifts

One record for each turn a volunteer does, past and on the rota ahead. The sample has 1,448 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Person | Points at People, required | The volunteer. |
| Role | Points at Volunteer Roles, required | The role done. |
| Shift Date | Date | The day of the shift. |
| Start Time | Short text | When it starts. |
| Hours | Number | How long it lasts, in hours. |

#### Funders

The trusts, foundations and councils the trust applies to for grants. The sample has 6 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Name | Short text, required, unique | The funder. |
| Kind | Short text | Charitable trust, community foundation, council or other. |
| Contact Name | Short text | The grants officer the trust deals with. |
| E-mail | Short text | Their e-mail address. |
| Phone | Short text | Their phone number. |

#### Grants

Each application for a grant, from the day it was sent to the report the funder wants at the end. The sample has 10 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Funder | Points at Funders, required | Who the application went to. |
| Purpose | Short text | What the money is for. |
| Applied On | Date | The day it was sent. |
| Amount Asked | Currency | How much was asked for. |
| Status | Short text | Applied (waiting for a decision), Awarded, Declined, or Completed once the report is sent. |
| Decided On | Date | The day the funder said yes or no. |
| Amount Awarded | Currency | How much was given; empty when declined. |
| Paid On | Date | The day the money arrived. |
| Period Ends | Date | The day the funded work should be finished. |
| Report Due | Date | The day the funder wants a report on how the money was spent. |
| Report Sent | Date | The day the report went; empty while it is still to write. |

#### Income

The cash book: one line for every payment received, whatever it was for, as the treasurer banks it. Fees, gifts, tickets, grants, tax claims and collections are all here, so income can be counted in one place. The sample has 3,486 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Received On | Date | The day the money came in. |
| Source | Short text | What kind of income: Membership, Donations, Events, Collections, Grants or Tax claims. |
| Description | Short text | What it was for, such as the appeal or event. |
| Amount | Currency | How much. |
| Method | Short text | How it was paid. |
| Reference | Short text | The record it came from, such as Gift 123 or Booking 45, so it can be traced. |

### Relationships

- Memberships.Person points at People.ID
- Memberships.Membership Type points at Membership Types.ID
- Donations.Person points at People.ID
- Donations.Campaign points at Campaigns.ID
- Donations.Tax Claim points at Tax Claims.ID
- Event Bookings.Event points at Events.ID
- Event Bookings.Person points at People.ID
- Volunteer Shifts.Person points at People.ID
- Volunteer Shifts.Role points at Volunteer Roles.ID
- Grants.Funder points at Funders.ID

### Queries

#### Current members

Everyone whose membership has started and not yet ended, and life members, by type. Use it to check who may vote at the AGM or use members' rates.

From Memberships, People, Membership Types. Shows Type, Member, Starts, Ends, E-mail, Phone.
Calculated: Member is `[First Name] & " " & [Last Name]`.
Only where `Days In >= 0` and `Days In >= 0` and `Days Left >= 0` and `Ends Is Null`.

#### Members by type

How many current members there are of each type and what they paid. A quick count for the committee.

From Memberships, Membership Types. Shows Type, Members, Fees Paid.
Only where `Days In >= 0` and `Days In >= 0` and `Days Left >= 0` and `Ends Is Null`.

#### Renewals due

Memberships ending in the next 30 days that have not been renewed, with addresses and the fee to ask for. Run it each week and send the reminders.

From Memberships, People, Membership Types. Shows Ends, Days Left, Member, Type, Renewal Fee, Address, City, Postcode, E-mail, Contact by E-mail.
Calculated: Days Left is `DateDiff("d", Date(), [Ends])`; Member is `[First Name] & " " & [Last Name]`.
Only where `Days Left Between 0 And 30` and `Renewed = No`.

#### Lapsed members

Members whose membership ended in the last year and who have not renewed. Worth a friendly call or e-mail; some simply forgot.

From Memberships, People, Membership Types. Shows Member, Type, Ended, Days Lapsed, E-mail, Phone, Contact by E-mail.
Calculated: Member is `[First Name] & " " & [Last Name]`; Days Lapsed is `DateDiff("d", [Ends], Date())`.
Only where `Days Lapsed Between 1 And 365` and `Renewed = No`.

#### Appeals against target

Each appeal and fund with its target, what has been given to it and how many gifts. Shows which appeals reached their target.

From Campaigns, Donations. Shows Campaign, Opens, Closes, Target, Given, Gifts.

#### Gifts to claim

Claimable gifts not yet put into a tax claim, with the giver's name, house and postcode and what the claim on each brings back. This is the next claim.

From Donations, People. Shows Gift Date, First Name, Last Name, Address, Postcode, Amount, Claim Value.
Calculated: Claim Value is `Round([Amount] * 0.25, 2)`.
Only where `Claimable = Yes` and `Tax Claim Is Null`.

#### Tax claims made

Each tax claim with the gifts in it, what was claimed and when the money came; a claim still awaited has no date received.

From Tax Claims, Donations. Shows Reference, Gifts Up To, Submitted On, Gifts, Gifts Total, Claimed, Received On, Received.

#### Top supporters

The people who gave most in the last 12 months, with the number of gifts. Use it for thank-you letters.

From People, Donations. Shows Supporter, City, Given, Gifts, Tax Declaration.
Calculated: Supporter is `[First Name] & " " & [Last Name]`.
Only where `Days Ago Between 0 And 365`.

#### Regular givers

Supporters with a monthly gift that arrived in the last 45 days, and the amount. A giver missing from it may have stopped their standing order.

From People, Donations. Shows Supporter, Monthly Gift, Last Gift, Tax Declaration.
Calculated: Supporter is `[First Name] & " " & [Last Name]`.
Only where `Regular = Yes` and `Days Ago Between 0 And 45`.

#### Volunteer hours by month

Hours given each month in each role, and their value at the role's hourly rate. Funders and annual reports often ask for this.

From Volunteer Shifts, Volunteer Roles. Shows Month, Role, Hours, Value.
Calculated: Month is `Format([Shift Date], "yyyy-mm")`; Value is `Round([Hours] * [Hour Value], 2)`.
Only where `Days Ago >= 0`.

#### Volunteers

Each volunteer's shifts and hours in the last 12 months and their last shift, most hours first.

From Volunteer Shifts, People. Shows Volunteer, Phone, Shifts, Hours, Last Shift.
Calculated: Volunteer is `[First Name] & " " & [Last Name]`.
Only where `Days Ago Between 0 And 365`.

#### Volunteer rota

Who is on which shift in the next 14 days, with a phone number, for the rota sheet on the noticeboard.

From Volunteer Shifts, Volunteer Roles, People. Shows Shift Date, Start Time, Role, Volunteer, Phone, Hours.
Calculated: Volunteer is `[First Name] & " " & [Last Name]`.
Only where `Days Ahead Between 0 And 14`.

#### Upcoming events

Events in the next 60 days with tickets booked so far against the capacity and the money taken.

From Events, Event Bookings. Shows Event Date, Event, Start Time, Venue, Capacity, Booked, Taken.
Only where `Days Ahead Between 0 And 60` and `Cancelled = No`.

#### Event attendance

Each event held, with tickets booked, how many came and the ticket income against the costs.

From Events, Event Bookings. Shows Event Date, Event, Capacity, Booked, Came, Ticket Income, Costs.
Calculated: Came is `IIf([Attended] = 1, [Tickets], 0)`.
Only where `Status = Held` and `Cancelled = No`.

#### Guests booked

Every booking for events in the next 60 days, by event, for the door list.

From Event Bookings, Events, People. Shows Event, Event Date, Guest, Tickets, Amount Paid, Phone.
Calculated: Event is `[Events].[Name] & ", " & Format([Event Date], "yyyy-mm-dd")`; Guest is `[First Name] & " " & [Last Name]`.
Only where `Days Ahead Between 0 And 60` and `Cancelled = No`.

#### Grant reports due

Grants awarded whose report to the funder is not yet sent, soonest first; a negative Days Left means the report is late.

From Grants, Funders. Shows Funder, Purpose, Amount Awarded, Period Ends, Report Due, Days Left, Contact Name.
Calculated: Days Left is `DateDiff("d", Date(), [Report Due])`.
Only where `Status = Awarded` and `Report Sent Is Null`.

#### Grants by funder

For each funder, how many applications were made, how much was asked and how much awarded.

From Grants, Funders. Shows Funder, Kind, Applications, Asked, Awarded.
Calculated: Awarded is `Nz([Amount Awarded])`.

#### Income by source by month

All income each month by source: membership, donations, events, collections, grants and tax claims.

From Income. Shows Month, Source, Received, Payments.
Calculated: Month is `Format([Received On], "yyyy-mm")`.

#### Income by source

Income in the last 12 months by source, largest first, for the treasurer's report.

From Income. Shows Source, Received, Payments.
Only where `Days Ago Between 0 And 365`.

#### Mailing list

Everyone who agreed to post, with their address, sorted by name, for the newsletter labels.

From People. Shows Last Name, First Name, Address, City, Postcode.
Only where `Contact by Post = Yes`.

### Reports

#### Membership list

The current members by type, with dates and contact details, and a count of each type. For the membership secretary.

From the query Current members, grouped by Type: Member, Starts, Ends, E-mail, Phone, with count of member.

#### Renewal reminders

Members due to renew in the next 30 days with their address and fee, to write or print the reminders from.

From the query Renewals due: Member, Address, City, Postcode, Type, Ends, Renewal Fee, with total of renewal fee.

#### Donations by campaign

Each appeal and fund with its target, the number of gifts and the total given. For the trustees' meeting.

From the query Appeals against target: Campaign, Opens, Closes, Target, Gifts, Given, with total of given, total of gifts.

#### Tax claim schedule

The gifts for the next tax claim with each giver's name, house and postcode, date and amount, and the total to claim. Keep it with the claim.

From the query Gifts to claim: First Name, Last Name, Address, Postcode, Gift Date, Amount, Claim Value, with total of amount, total of claim value.

#### Volunteer hours

Volunteer hours and their value by month and role, with totals. For the annual report and grant reports.

From the query Volunteer hours by month, grouped by Month: Role, Hours, Value, with total of hours, total of value.

#### Event guest list

The guests booked for each event in the next 60 days with their tickets, for the door.

From the query Guests booked, grouped by Event: Guest, Tickets, Amount Paid, Phone, with total of tickets, total of amount paid.

#### Income by month

Income by source for each month with the month's total. For the treasurer.

From the query Income by source by month, grouped by Month: Source, Payments, Received, with total of received.

#### Grant reports to write

Grant reports still to write, with the amount and the date each is due.

From the query Grant reports due: Funder, Purpose, Amount Awarded, Report Due, Days Left, with total of amount awarded.

### Starting empty

Start with the lookup tables: Membership Types (your fees), Campaigns (at least a General Fund for gifts given to nothing in particular), Volunteer Roles and Funders. Then enter People as you meet them, with what they have agreed to be contacted by. After that, record Memberships, Donations, Events and their Bookings, Volunteer Shifts and Grants as they happen, and write each payment received in Income too. Make a Tax Claim when you send one and set each gift's Tax Claim to it.

### Making it yours

- If your country has no scheme for claiming tax on gifts, leave Tax Declaration empty and delete the Tax Claims table, the Gifts to claim query and the Claim schedule report.
- If your scheme pays a different share than a quarter, change the 0.25 in the Claim Value column of the Gifts to claim query.
- For a club with monthly subscriptions, add a Monthly membership type with Months set to 1, and the renewal queries work as before.
- Add a Spending table (date, kind, description, amount) to see income against costs each month.
- Add an Event field to Volunteer Shifts, pointing at Events, to see who stewarded each event.
- For a parent teacher association, rename Memberships to Families and add a field for the children's classes.
- Add a Gift in Memory Of field to Donations if people often give in memory of someone.
- Add a Safeguarding Check date to People, and a query of volunteers whose check is over three years old.

## Granite and Grain Woodworks

**Workshop or maker.** Materials, bills of materials, work orders, makers' hours, orders and deliveries for a small furniture workshop.

Granite and Grain Woodworks is an invented workshop of five makers in a unit on an industrial estate. They build dining tables, chairs, stools, benches, shelves, bookcases and chopping boards from oak, ash, walnut and beech, and make kitchen doors, drawer fronts and worktops to measure. About ten gift and homeware shops buy boards, coasters and shelves at a trade discount, interior designers order kitchens and furniture for their clients, and members of the public buy from the workshop's showroom and website.

Every product has a bill of materials: the timber, sheet, fittings, finish and packing that one of it takes. Small goods, chairs and shelves are made in batches to keep a stock on the shelf. Tables, bookcases, worktops and kitchen doors are made to order, so each order line of them gets a work order of its own. Planning a work order sets its materials aside, starting it takes them from stock, and the makers book their hours to it each day until it is finished.

The records cover fifteen months up to the morning the file is made, with orders already taken for the weeks ahead. Shops order more boards before Christmas, kitchens slow in August, the workshop closes between Christmas and New Year, and timber prices went up part way through. Some deliveries are late, most of them in January after the break or when a maker was off sick, and a few trade customers pay late. The buyer orders twice a week for the work starting in the next fortnight, so a job further off can show as a shortage until it comes closer.

Use it to see how a bill of materials, stock and production fit together, or start from it empty for a workshop of your own.

It suits, with small changes: A cabinet maker or joiner; A bakery's production side; A print shop; A sign maker; A soap or candle maker; A garment or textile maker; An electronics assembler; A metal fabricator.

### Tables

#### Suppliers

The merchants the workshop buys its materials from. The sample has 6 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Name | Short text, required | The supplier's name. |
| Supplies | Short text | What you buy from them, such as Timber or Fittings and fixings. |
| Contact | Short text | The person you order from. |
| Phone | Short text | Their phone number. |
| E-mail | Short text | Where orders go. |
| City | Short text | Where they are. |
| Lead Days | Whole number | How many days an order usually takes to arrive, used to judge when to order. |
| Minimum Order | Currency | The smallest order they deliver without a charge. The buyer waits until enough is needed from them, unless work would be held up or the order has waited a fortnight. |

#### Materials

Everything that goes into a product: timber, sheet materials, fittings, fixings, finishes and packing. One record per thing you buy and keep. The sample has 27 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Code | Short text, unique | A short code for the material, used on the bench and on purchase orders. Each code is used once. |
| Name | Short text, required | What it is, with its size. |
| Kind | Short text | Timber, Sheet, Fittings, Fixings, Finishes or Packing. |
| Unit | Short text | How it is counted: metre, sheet, each, pair, box, litre or roll. |
| Unit Cost | Currency | What one unit costs now. Change it when the supplier's price changes, then work out each product's Material Cost again. |
| Supplier | Points at Suppliers | Who you usually buy it from. |
| In Stock | Number | How much is in the workshop now. |
| On Order | Number | How much has been ordered and not yet arrived. |
| Allocated | Number | How much is set aside for work orders that are planned but not started. It goes up when a work order is planned and down when work starts and the material is taken. |
| Reorder Level | Number | When the free stock (in stock, plus on order, less allocated) falls to this or below, the material goes on the reorder list. |
| Reorder Quantity | Number | How much you usually order at a time. |

#### Material Purchases

Each order placed with a supplier. The sample has 266 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Supplier | Points at Suppliers, required | Who it was ordered from. |
| Ordered On | Date | The day it was placed. |
| Expected | Date | When the supplier said it would come. |
| Received On | Date | The day it arrived; empty until then. |
| Status | Short text | Ordered until it arrives, then Received. |

#### Purchase Lines

What was ordered on each purchase. The sample has 457 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Purchase | Points at Material Purchases, required | The purchase the line belongs to. |
| Material | Points at Materials, required | What was ordered. |
| Quantity | Number | How much, in the material's unit. |
| Unit Cost | Currency | The price of one unit on this order, kept so a later price change does not alter it. |

#### Material Movements

Every change to a material's stock, so the In Stock figure can be traced. The sample has 3,997 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Material | Points at Materials, required | The material. |
| When | Date | The day it happened. |
| Kind | Short text | Received (a delivery came in), Used (taken for a work order) or Counted (the opening count, or a correction after a stock count). |
| Quantity | Number | How much came in (above nought) or went out (below nought). |
| Reference | Short text | The purchase or work order it belongs to, such as PO 12 or WO 340. |

#### Products

What the workshop makes and sells, one record per product. The sample has 24 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Code | Short text, unique | A short product code. Each code is used once. |
| Name | Short text, required | The product as it shows on orders and invoices. |
| Category | Short text | Tables, Seating, Storage, Small goods or Kitchen. |
| Price | Currency | The price to the public before any trade discount. |
| Material Cost | Currency | The cost of one product's materials from its bill of materials, worked out at today's unit costs. The Costed bill of materials query shows how it is made up. |
| Labour Minutes | Whole number | How long one takes to make, in minutes, by a maker working at a normal pace. |
| Made to Order | Yes/No | Ticked for products made for each order (tables, bookcases, worktops, kitchen doors); left clear for products kept on the shelf. |
| In Stock | Whole number | Finished pieces on the shelf, including made-to-order pieces finished and waiting for delivery. |
| Stock Level | Whole number | How many to keep on the shelf beyond what is already ordered. A batch is planned when stock and work in hand fall short of it. |
| Batch Size | Whole number | How many are made at once. |
| Active | Yes/No | Untick a product you no longer make; it stays in the old orders. |

#### Bill of Materials

What one of each product is made from: one record per material in it. The sample has 150 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Product | Points at Products, required | The product. |
| Material | Points at Materials, required | A material it takes. |
| Quantity | Number | How much of the material one product uses, in the material's unit, waste included. |

#### Makers

The people in the workshop who make things and book their hours. The sample has 5 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| First Name | Short text | Their first name. |
| Last Name | Short text, required | Their family name. |
| Role | Short text | Workshop lead, cabinet maker, chair maker or apprentice. |
| Hourly Cost | Currency | What an hour of their time costs the workshop, with holiday pay and insurance. |
| Phone | Short text | A number to reach them. |
| E-mail | Short text | Their work address. |
| Started | Date | The day they joined. |

#### Customers

The shops, designers and people who buy from the workshop. The sample has 197 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Name | Short text, required | The business, or the person's name for a member of the public. |
| Kind | Short text | Shop, Designer or Public. |
| Contact | Short text | The person you deal with at a business. |
| Phone | Short text | Their phone number. |
| E-mail | Short text | Where invoices go. |
| Address | Short text | The street address for deliveries. |
| City | Short text | The town or city. |
| Postcode | Short text | The postcode. |
| Payment Days | Whole number | Days allowed to pay after delivery: 30 for shops, 14 or 30 for designers, nought for the public, who pay on delivery. |
| Trade Discount | Number | The discount on the public price, as a fraction: 0.3 is thirty percent off. |

#### Sales Orders

Each order a customer places. The sample has 451 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Customer | Points at Customers, required | Who ordered. |
| Order Date | Date | The day the order was taken. |
| Delivery Due | Date | The day you promised to deliver. |
| Delivered On | Date | The day it was delivered; empty until then. |
| Status | Short text | Open, Delivered or Cancelled. |
| Delivery Fee | Currency | What was charged for delivery; nought when collected. |
| Invoice Amount | Currency | The invoice total, goods after discount plus delivery, set on delivery. |
| Invoice Due | Date | The day payment is due, the delivery day plus the customer's payment days. |
| Paid On | Date | The day it was paid; empty while it is owed. |
| Notes | Long text | Anything to remember about the order. |

#### Sales Lines

The products on each order. The sample has 1,036 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Order | Points at Sales Orders, required | The order the line belongs to. |
| Product | Points at Products, required | What was ordered. |
| Quantity | Whole number | How many. |
| Unit Price | Currency | The price of one at the time, kept so a later price change does not alter old orders. |
| Discount | Number | The discount given, as a fraction: 0.1 is ten percent off. |

#### Time Sheets

The hours each maker booked to each work order, one record per maker, work order and day. The sample has 1,593 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Maker | Points at Makers, required | Who did the work. |
| Work Order | Points at Work Orders, required | What they worked on. |
| Worked On | Date | The day. |
| Hours | Number | Hours worked on it that day, in quarter hours. |

#### Work Orders

A job to make a quantity of one product, either for a customer's order or to fill the shelf. The sample has 619 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Product | Points at Products, required | What is to be made. |
| Quantity | Whole number | How many. |
| Maker | Points at Makers | The maker the job is given to. |
| Sales Order | Points at Sales Orders | The order it is made for; empty for a batch made for stock. |
| Planned Start | Date | When work is planned to begin. |
| Due Date | Date | When it must be finished: two working days before the delivery for an order. |
| Started | Date | The day work began and the materials were taken. |
| Finished | Date | The day the last piece was finished. |
| Status | Short text | Planned, In progress, Done or Cancelled. |

#### Stock Movements

Every change to the stock of finished goods. The sample has 1,587 records.

| Field | Type | What it holds |
|---|---|---|
| ID | AutoNumber | The number Data Files gives each record. Other tables point at a record by it. |
| Product | Points at Products, required | The product. |
| When | Date | The day it happened. |
| Kind | Short text | Opening (the count the records start from), Made (a work order finished), Delivered (sent to a customer) or Adjusted (damaged or found at a count). |
| Quantity | Whole number | How many came in (above nought) or went out (below nought). |
| Reference | Short text | The work order or sales order it belongs to. |

### Relationships

- Materials.Supplier points at Suppliers.ID
- Material Purchases.Supplier points at Suppliers.ID
- Purchase Lines.Purchase points at Material Purchases.ID
- Purchase Lines.Material points at Materials.ID
- Material Movements.Material points at Materials.ID
- Bill of Materials.Product points at Products.ID
- Bill of Materials.Material points at Materials.ID
- Sales Orders.Customer points at Customers.ID
- Sales Lines.Order points at Sales Orders.ID
- Sales Lines.Product points at Products.ID
- Time Sheets.Maker points at Makers.ID
- Time Sheets.Work Order points at Work Orders.ID
- Work Orders.Product points at Products.ID
- Work Orders.Maker points at Makers.ID
- Work Orders.Sales Order points at Sales Orders.ID
- Stock Movements.Product points at Products.ID

### Queries

#### Costed bill of materials

Every material in every product, with how much it takes, the unit cost and the line cost. Use it to check a price or to see what a price rise in one timber does to each product.

From Products, Bill of Materials, Materials. Shows Product Code, Product, Material, Quantity, Unit, Unit Cost, Line Cost.
Calculated: Line Cost is `Round([Bill of Materials].[Quantity] * [Unit Cost], 2)`.
Only where `Active = Yes`.

#### Product margins

Each product's price less its material cost and its labour, with the margin as a percentage of the price, the best first. Labour is costed at 21 an hour; change the number in the query to your own rate.

From Products. Shows Code, Product, Category, Price, Material Cost, Labour Cost, Margin, Margin %.
Calculated: Labour Cost is `Round([Labour Minutes] / 60 * 21, 2)`; Margin is `Round([Price] - [Material Cost] - [Labour Minutes] / 60 * 21, 2)`; Margin % is `Round(100 * ([Price] - [Material Cost] - [Labour Minutes] / 60 * 21) / [Price], 1)`.
Only where `Active = Yes`.

#### Materials for planned work

What the planned work orders need of each material, from their bills of materials, beside what is allocated and in stock. The Needed column is how the Allocated figure is made up.

From Materials, Bill of Materials, Products, Work Orders. Shows Code, Material, Unit, Needed, Allocated, In Stock.
Calculated: Needed is `Round([Work Orders].[Quantity] * [Bill of Materials].[Quantity], 2)`.
Only where `Status = Planned`.

#### Material shortages

Materials where the planned work needs more than is in stock and on order, with how much is short. Buy these before the work that needs them is due to start, or it will wait.

From Materials, Suppliers. Shows Supplier, Code, Material, Unit, Allocated, In Stock, On Order, Short By.
Calculated: Short By is `Round([Allocated] - [In Stock] - [On Order], 2)`.
Only where `Short By > 0`.

#### Materials to reorder

Materials whose free stock (in stock, plus on order, less allocated to planned work) is at or below the reorder level, by supplier, with the minimum order for that supplier and how much you usually order. Free stock below nought is a shortage.

From Materials, Suppliers. Shows Supplier, Minimum Order, Code, Material, Unit, Free Stock, Reorder Level, Reorder Quantity.
Calculated: Free Stock is `Round([In Stock] + [On Order] - [Allocated], 2)`.
Only where `Above Reorder Level <= 0`.

#### Purchases on order

Orders placed with suppliers that have not arrived, with when each is expected and what it is worth.

From Material Purchases, Suppliers, Purchase Lines. Shows Purchase, Supplier, Ordered On, Expected, Value.
Calculated: Value is `Round([Quantity] * [Unit Cost], 2)`.
Only where `Status = Ordered`.

#### Open work orders

Work planned or in progress, the soonest due first, with the maker, the order it is for and the days left. Days left below nought means it is late.

From Work Orders, Products, Makers. Shows Work Order, Product, Quantity, Maker, For Order, Planned Start, Due Date, Started, Status, Days Left.
Calculated: Maker is `[First Name] & " " & [Last Name]`; Days Left is `DateDiff("d", Date(), [Due Date])`.
Only where `Status In ("Planned", "In progress")`.

#### Work order costs

Work orders finished in the last 90 days, with the hours booked against the standard hours, the labour cost at each maker's rate and the material cost.

From Time Sheets, Makers, Work Orders, Products. Shows Work Order, Product, Quantity, Finished, Standard Hours, Hours Booked, Labour Cost, Material Cost.
Calculated: Standard Hours is `Round([Work Orders].[Quantity] * [Labour Minutes] / 60, 2)`; Labour Cost is `Round([Hours] * [Hourly Cost], 2)`; Material Cost is `Round([Work Orders].[Quantity] * [Material Cost], 2)`.
Only where `Status = Done` and `Days Since Between 0 And 90`.

#### Maker hours by month

The hours each maker booked to work orders in each month, and what they cost.

From Time Sheets, Makers. Shows Maker, Month, Hours, Cost.
Calculated: Maker is `[First Name] & " " & [Last Name]`; Month is `Format([Worked On], "yyyy-mm")`; Cost is `Round([Hours] * [Hourly Cost], 2)`.

#### Maker hours, last 4 weeks

The hours each maker booked over the last four weeks, to see who is busy and who has time.

From Time Sheets, Makers. Shows Maker, Role, Hours, Cost.
Calculated: Maker is `[First Name] & " " & [Last Name]`; Cost is `Round([Hours] * [Hourly Cost], 2)`.
Only where `Days Ago Between 0 And 27`.

#### Orders to deliver

Open orders due in the next 14 days, and any already late, the soonest first, with what the goods come to.

From Sales Orders, Customers, Sales Lines. Shows Order, Customer, Kind, Order Date, Delivery Due, Days Left, Goods.
Calculated: Days Left is `DateDiff("d", Date(), [Delivery Due])`; Goods is `Round([Quantity] * [Unit Price] * (1 - Nz([Discount])), 2)`.
Only where `Days Left <= 14` and `Status = Open`.

#### Late deliveries

Orders delivered after the promised day in the last year, with how many days late.

From Sales Orders, Customers. Shows Order, Customer, Kind, Delivery Due, Delivered On, Days Late.
Calculated: Days Late is `DateDiff("d", [Delivery Due], [Delivered On])`.
Only where `Days Late > 0` and `Days Ago Between 0 And 365`.

#### Unpaid invoices

Orders delivered and not yet paid, the oldest due date first, with the days overdue. Below nought means it is not due yet.

From Sales Orders, Customers. Shows Order, Customer, Kind, Delivered On, Invoice Due, Invoice Amount, Days Overdue.
Calculated: Days Overdue is `DateDiff("d", [Invoice Due], Date())`.
Only where `Status = Delivered` and `Paid On Is Null`.

#### Finished goods in stock

Products on the shelf, with what they are worth at cost and at price.

From Products. Shows Category, Code, Product, In Stock, Stock Level, Value at Cost, Value at Price.
Calculated: Value at Cost is `Round([In Stock] * ([Material Cost] + [Labour Minutes] / 60 * 21), 2)`; Value at Price is `Round([In Stock] * [Price], 2)`.
Only where `In Stock > 0`.

#### Sales by month

Goods delivered in each month, after discounts and without delivery fees.

From Sales Orders, Sales Lines. Shows Month, Items, Sales.
Calculated: Month is `Format([Delivered On], "yyyy-mm")`; Sales is `Round([Quantity] * [Unit Price] * (1 - Nz([Discount])), 2)`.
Only where `Status = Delivered`.

#### Sales by product

What each product has sold, the biggest seller first.

From Sales Lines, Products, Sales Orders. Shows Category, Product, Items, Sales.
Calculated: Sales is `Round([Quantity] * [Unit Price] * (1 - Nz([Discount])), 2)`.
Only where `Status = Delivered`.

#### Sales by customer

What each customer has bought, the biggest first, with how many order lines.

From Customers, Sales Orders, Sales Lines. Shows Customer, Kind, Order Lines, Sales.
Calculated: Sales is `Round([Quantity] * [Unit Price] * (1 - Nz([Discount])), 2)`.
Only where `Status = Delivered`.

### Reports

#### Costed bill of materials

A page for each product listing its materials and their cost, with the total material cost.

From the query Costed bill of materials, grouped by Product: Material, Quantity, Unit, Unit Cost, Line Cost, with total of line cost.

#### Product margins

Every product with its price, costs and margin, for setting prices and deciding what to push.

From the query Product margins: Code, Product, Price, Material Cost, Labour Cost, Margin, Margin %.

#### Production plan

The open work orders for each maker in the order they are planned to start, for the Monday meeting and the workshop wall.

From the query Open work orders, grouped by Maker: Work Order, Product, Quantity, Status, Planned Start, Due Date, Days Left, with count of work order.

#### Shortages for buying

The materials short for planned work, by supplier, for the buyer to order first.

From the query Material shortages, grouped by Supplier: Code, Material, Unit, Allocated, In Stock, On Order, Short By.

#### Reorder list

Materials at or below their reorder level, by supplier, with how much to order.

From the query Materials to reorder, grouped by Supplier: Code, Material, Unit, Free Stock, Reorder Level, Reorder Quantity.

#### Delivery schedule

Orders to deliver in the next two weeks and those late, for planning the van.

From the query Orders to deliver: Order, Customer, Kind, Order Date, Delivery Due, Days Left, Goods, with total of goods.

#### Monthly sales

Sales for each month, with a total.

From the query Sales by month: Month, Items, Sales, with total of items, total of sales.

#### Invoices due

What each customer owes, with the due dates and days overdue, for chasing payment.

From the query Unpaid invoices, grouped by Customer: Order, Delivered On, Invoice Due, Days Overdue, Invoice Amount, with total of invoice amount.

#### Hours by maker

Each maker's booked hours and their cost month by month, with totals.

From the query Maker hours by month, grouped by Maker: Month, Hours, Cost, with total of hours, total of cost.

### Starting empty

Start with Suppliers, then Materials with their units, costs and reorder levels, then Products, and then the Bill of Materials for each product. Fill in each product's Material Cost from the Costed bill of materials query. Add your Makers and Customers. After that, Sales Orders and their lines come in as you trade, Work Orders are raised for them, and the makers book Time Sheets as they work. Keep In Stock, On Order and Allocated up to date as deliveries arrive and work starts, or count once a month and correct them.

### Making it yours

- For a bakery or candle maker, change the units to kilogram and gram, make Quantity in Bill of Materials hold small amounts, and add a Batch Yield field to Products for how many one batch makes.
- Record finishes and options: add Finish and Size fields to Sales Lines for made-to-order pieces, and copy them to the work order.
- Track offcuts and waste: add a Waste Percent field to Bill of Materials and use it in the Costed bill of materials query.
- Add VAT or sales tax: add a Tax Rate field to Products and a calculated column in the sales queries for the tax in each line.
- Take deposits on commissions: add Deposit and Deposit Paid On fields to Sales Orders, and a query for orders waiting for one.
- Plan by machine as well as maker: add a Machines table and a Machine field to Work Orders, then group the production plan by it.
- Keep batch or lot numbers for traceability: add a Batch field to Material Movements and Stock Movements.
