// The documentation for Data Files' ready-made databases (templates.js), in
// the same shape as each sample business's `about`: what each one is for,
// every table and field, its query and report, a few everyday tasks done
// step by step, things worth knowing, questions and ways to change it. The
// templates come empty, so there are no invented records to describe here,
// and no print templates yet.

export const TEMPLATE_ABOUT = {
  contacts: {
    name: 'Contacts',
    kind: 'Address book',
    line: 'People you know or deal with, how to reach them, and a phone list to print.',
    summary: [
      'Contacts is an address book for anyone who keeps a list of people: a small office, a club secretary, a tradesman with regular customers, or a household. It starts empty and holds one record for each person.',
      'Each record has the person\'s name, the company they work for and their job, their e-mail, phone and mobile numbers, a postal address, a birthday and room for notes. Only the last name has to be filled in, so a record can start with very little and grow.',
      'There is a single table, so nothing points anywhere else. The company is typed on each person rather than chosen from a list, which keeps it simple; the Customers and sales template keeps companies in a table of their own when you need that.',
    ],
    suits: ['A small office keeping its business contacts', 'A club or society secretary', 'A self-employed tradesman or cleaner with regular customers', 'A family address book with birthdays', 'A school or nursery keeping parents\' details', 'An event organiser keeping suppliers and helpers'],
    tables: {
      Contacts: { about: 'One record for each person you keep in touch with.', fields: {
        'First Name': 'Their first or given name.',
        'Last Name': 'Their family name. It must be filled in, and the list and the phone list are sorted by it.',
        Company: 'The company or organisation they work for, typed as you would write it.',
        'Job Title': 'Their job, such as Office Manager or Head Teacher.',
        'E-mail': 'Their e-mail address.',
        Phone: 'Their landline or work number.',
        Mobile: 'Their mobile number.',
        Address: 'The street address: the house number and street, and a second line if needed.',
        City: 'The town or city.',
        Postcode: 'The postcode or ZIP code.',
        Country: 'The country, useful when your contacts are in more than one.',
        Birthday: 'The day they were born. The year can be left as a guess if you only want the day.',
        Notes: 'Anything worth remembering: how you met, a partner\'s name, the best time to call.',
      } },
    },
    queries: {
      'Contact list': 'Every contact with their company, e-mail and phone, sorted by last name and then first name. There are no criteria, so it shows everyone; use it to look someone up or copy the list elsewhere.',
    },
    reports: {
      'Phone list': 'A printed list of every contact by last name, with their phone, mobile and e-mail, to pin up by the phone or keep in a folder.',
    },
    documents: {},
    tutorial: [
      { title: 'Add a contact', steps: [
        'Open the Contacts form from the Forms list and press New Record.',
        'Type at least the Last Name, then whatever else you know: First Name, Company, Phone, Mobile and E-mail.',
        'Press Add Record. The new contact is kept and the form shows it.',
      ] },
      { title: 'Change someone\'s details', steps: [
        'Open the Contacts form and move to the person with the Next record and Previous record arrows.',
        'Click in the box to change, such as Phone, type the new number and press Enter or click in another box. The change is kept as soon as you leave the box.',
        'For many changes at once, open the Contacts table from the Tables list instead: every contact is a row, and a cell you type in is kept the same way.',
      ] },
      { title: 'Look someone up', steps: [
        'Run the query Contact list from the Queries list. It shows everyone by last name.',
        'To see only the people at one company, open the query\'s design and type the company\'s name in the criteria under Company, in quotes: "Hill Street Dental".',
        'Run it again to see the shorter list. Clear the criteria to go back to everyone.',
      ] },
      { title: 'Print the phone list', steps: [
        'Open the report Phone list from the Reports list.',
        'Check the order in Sort by: Last Name keeps it alphabetical.',
        'Press Print, or Save as PDF to e-mail it or keep a copy.',
      ] },
      { title: 'Remove a contact', steps: [
        'Open the Contacts form, move to the person and press Delete.',
        'Before you delete, check it is the right person: there is no undo for a deleted record other than a copy made with File, Save As.',
        'To remove several, open the Contacts table from the Tables list, tick the box at the start of each row and press Delete.',
      ] },
    ],
    notes: [
      'Only Last Name must be filled in. For a company with no named person, type the company\'s name in Last Name and leave First Name empty.',
      'A change in a form or in the table is kept when you leave the box: there is no Save button to press for each record.',
      'There are no pointing fields in this template, so Company is plain text. Type a company the same way each time ("Hill Street Dental", not "Hill St Dental") so a query can find all its people.',
      'The whole address book is one file on this computer. Keep a copy elsewhere with File, Save As, on a memory stick or a cloud folder, and do it again from time to time.',
    ],
    faq: [
      { q: 'How do I sort the list by company instead of name?', a: 'In the report Phone list, choose Company in Sort by. In the query Contact list, open its design and move the sort from Last Name to Company.' },
      { q: 'Can I keep more than one e-mail address for a person?', a: 'Not as it comes. Add a second field to the Contacts table in its design, such as Work E-mail, and it appears on the form.' },
      { q: 'Why does the form refuse to add a contact?', a: 'Last Name is empty. Type something in it, even a company\'s name, and press Add Record again.' },
      { q: 'Can I bring in contacts I already have in a spreadsheet?', a: 'Save the spreadsheet as a CSV file and bring it in with Import. It comes in as a table of its own, which you can keep beside Contacts or use in its place.' },
      { q: 'How do I see whose birthday is coming up?', a: 'Make a new query on Contacts with First Name, Last Name and Birthday, sorted by Birthday. The template has no birthday query of its own.' },
      { q: 'Is my address book sent anywhere?', a: 'No. It is a file on this computer and stays there unless you copy or send it yourself.' },
    ],
    empty: 'There is only one table, so start straight away: open the Contacts form and add the people you call most, then fill in the rest as you come to them. If your contacts are already in a spreadsheet, you can save it as a CSV file and bring it in with Import as a table of its own.',
    adapt: [
      'Group your contacts: add a Group field (Family, Work, Club, Supplier) and a query that chooses one group in its criteria.',
      'Keep companies once: add a Companies table and change Company into a field that points at it, so the form gives a list of names to choose from.',
      'Send cards: add a Christmas Card field (Yes/No) and a query of everyone ticked, with their address, to print labels from.',
      'Record when you last spoke: add a Last Contacted date and sort the Contact list query by it to see who you have not called for a while.',
      'Add the social side: a Website field, or one for a messaging name, for people you reach that way.',
    ],
  },

  inventory: {
    name: 'Inventory',
    kind: 'Stock control',
    line: 'Products, the suppliers you buy them from, every stock movement and what is running low.',
    summary: [
      'Inventory is for anyone who holds goods and needs to know what is on the shelf and when to buy more: a small shop, a workshop with spare parts, a stores cupboard or an online seller working from a spare room.',
      'It keeps your suppliers, the products you stock with their price, the units in stock and the level at which you reorder, and a record of each time stock comes in or goes out. A query lists what is getting low with the supplier\'s phone number beside it, and a report lists every product under its supplier.',
      'Each product points at the supplier you buy it from, and each stock movement points at a product. One supplier has many products, and one product has many movements.',
    ],
    suits: ['A small shop or market stall', 'A workshop or garage keeping spare parts', 'An office or school stores cupboard', 'An online seller packing orders at home', 'A cafe or bar keeping dry goods and drinks', 'A charity shop or food bank'],
    tables: {
      Suppliers: { about: 'The businesses you buy from, one record each.', fields: {
        Name: 'The supplier\'s name. It must be filled in, and it is what the Supplier list on a product shows.',
        Contact: 'The person you deal with there.',
        'E-mail': 'Where to send an order.',
        Phone: 'The number to call when you need to reorder.',
        Address: 'Their address, for returns or collecting goods.',
      } },
      Products: { about: 'What you keep in stock, one record for each line.', fields: {
        Name: 'The product as you know it. It must be filled in, and it is what the Product list on a stock movement shows.',
        SKU: 'Your own code for the product, or the barcode number. Each SKU can be used once only.',
        Category: 'The group it belongs to, such as Drinks, Stationery or Brake Parts, typed the same way each time.',
        Supplier: 'Who you buy it from, chosen from your Suppliers.',
        'Unit Price': 'What one unit sells for, or what it costs you if you only track stock.',
        'Units in Stock': 'How many are on the shelf now. A new product starts at nought.',
        'Reorder Level': 'The number at which you mean to order more. It starts at 5.',
        Discontinued: 'Ticked for a product you no longer stock. It drops out of the Low stock query but keeps its history.',
      } },
      'Stock Movements': { about: 'Each time stock comes in or goes out: a delivery, a sale, a breakage or a stock count.', fields: {
        Product: 'The product that moved, chosen from your Products. It must be filled in.',
        When: 'The day it moved.',
        Quantity: 'How many: above nought for goods coming in, below nought for goods going out (-3 for three sold).',
        Note: 'Why it moved: Delivery, Sold, Damaged, Stock count, or a reference such as an order number.',
      } },
    },
    queries: {
      'Low stock': 'Products with fewer than 10 units in stock, by name, with the reorder level, the supplier and the supplier\'s phone. Discontinued products are left out. Run it before you place an order.',
    },
    reports: {
      'Products by supplier': 'Every product under the supplier you buy it from, by name, with its SKU, category, price and units in stock, and the units in stock added up for each supplier and for the whole. Use it as a stock sheet or to check what you buy from whom.',
    },
    documents: {},
    tutorial: [
      { title: 'Add a supplier and a product', steps: [
        'Open the Suppliers form, press New Record, type the Name, Contact, E-mail and Phone, and press Add Record.',
        'Open the Products form and press New Record. Type the Name, an SKU and a Category, and choose the supplier in the Supplier list.',
        'Type the Unit Price, the Units in Stock you have now and a Reorder Level that suits how quickly it sells, then press Add Record.',
      ] },
      { title: 'Book in a delivery', steps: [
        'Open the Stock Movements form and press New Record.',
        'Choose the product in Product, type the day in When and the number received in Quantity (24 for twenty-four). Type Delivery in Note and press Add Record.',
        'Open the Products form, move to the same product and add the same number to Units in Stock.',
      ] },
      { title: 'Record goods going out', steps: [
        'Open the Stock Movements form, press New Record and choose the Product.',
        'Type the day in When and the number with a minus in Quantity (-2 for two sold). Type Sold or Damaged in Note and press Add Record.',
        'On the Products form, take the same number off Units in Stock.',
      ] },
      { title: 'See what to reorder', steps: [
        'Run the query Low stock from the Queries list.',
        'Each line shows the product, how many are left, its reorder level, the supplier and the supplier\'s Phone, ready to call.',
        'To order sooner or later, open the query\'s design and change the 10 in the criteria under Units in Stock.',
      ] },
      { title: 'Print a stock sheet', steps: [
        'Open the report Products by supplier from the Reports list.',
        'It lists each supplier\'s products with the units in stock and a total for each supplier.',
        'Press Print to take round the shelves for a stock count, or Save as PDF.',
      ] },
      { title: 'Put the stock right after a count', steps: [
        'Count the shelf. If it differs from Units in Stock, open the Stock Movements form and add a movement for the difference, above or below nought, with Stock count in Note.',
        'Then type the counted number in Units in Stock on the Products form.',
        'Run the query Low stock afterwards: the count may show something lower than you thought.',
      ] },
    ],
    notes: [
      'Units in Stock is a number you keep up to date yourself: adding a stock movement does not change it. Change both together, and count the stock now and then to check them.',
      'Supplier on a product, and Product on a stock movement, are lists of names on the form: you choose from what is already in the other table. Add the supplier before its products, and the product before its movements. In the table\'s rows the same field shows the record\'s number instead.',
      'The Low stock query uses a fixed 10, not each product\'s Reorder Level, which it shows beside it for you to compare. Change the 10 in its design to suit your stock.',
      'A supplier with products cannot be deleted while those products point at it, and a product with movements cannot be deleted while they point at it. Tick Discontinued on a product you no longer keep instead.',
      'Keep a copy of the file with File, Save As before a big change such as a year-end stock count.',
    ],
    faq: [
      { q: 'Why did Units in Stock not change when I added a movement?', a: 'The template keeps the two apart: Stock Movements is the history, Units in Stock is today\'s number. Change Units in Stock on the Products form as well.' },
      { q: 'Can a product come from more than one supplier?', a: 'Not as it comes: each product points at one supplier. Choose the one you use most, or add a Second Supplier field that points at Suppliers too.' },
      { q: 'What should I put in Unit Price?', a: 'Whatever you need to know most often. A shop usually keeps the selling price; a stores cupboard the cost. To keep both, add a Cost Price field.' },
      { q: 'Why can I not use the same SKU twice?', a: 'SKU is unique, so each code points at one product. If two products share a barcode, add a letter to one of them.' },
      { q: 'How do I see every movement for one product?', a: 'Open the Products form at that product: the stock movements that point at it are listed below its fields.' },
      { q: 'A discontinued product still shows on the stock sheet. Is that right?', a: 'Yes. The report Products by supplier lists every product. Only the Low stock query leaves discontinued lines out.' },
    ],
    empty: 'Enter your Suppliers first, then your Products, choosing each one\'s supplier from the list and typing what is on the shelf now in Units in Stock. From then on add a stock movement each time goods come in or go out. If you want the history to match, add one movement per product for the stock you start with, with Opening stock in Note.',
    adapt: [
      'Keep a buying price as well as a selling price: add a Cost Price field to Products and a query with a calculated column for the stock\'s value at cost.',
      'Hold stock in more than one place: add a Locations table, a Location field on Stock Movements that points at it, and a query of movements by location.',
      'Say what kind of movement it is: add a Kind field to Stock Movements (Received, Sold, Damaged, Counted) instead of writing it in Note.',
      'Record purchase orders: add an Orders table that points at Suppliers and an Order Lines table that points at Orders and Products.',
      'Use each product\'s own reorder level: change the Low stock criteria from a fixed 10 to the numbers that suit each group of products, or make one query per Category.',
    ],
  },

  membership: {
    name: 'Membership',
    kind: 'Clubs and societies',
    line: 'Members, what they have paid and the renewals coming due.',
    summary: [
      'Membership is for a club, society or association that has members who join, pay and renew: a sports club, a choir, an allotment society, a residents\' association or a small professional body.',
      'It keeps each member\'s name and how to reach them, the type of membership they hold, when they joined, when their renewal is due and whether they are still active. Every payment a member makes is kept as its own record, so you can see what each person has paid over the years.',
      'Each payment points at the member who made it. One member has many payments. A query lists the active members by their renewal date, and a report adds up what each member has paid.',
    ],
    suits: ['A sports or social club', 'A choir, band or drama group', 'An allotment or gardening society', 'A residents\' or parents\' association', 'A small trade or professional body', 'A gym or dance class with members who pay by the year'],
    tables: {
      Members: { about: 'One record for each member, past or present.', fields: {
        'First Name': 'Their first or given name.',
        'Last Name': 'Their family name. It must be filled in.',
        'E-mail': 'Where to send renewal reminders and news.',
        Phone: 'Their phone number.',
        Address: 'Their postal address, for anything you send by post.',
        'Membership Type': 'The kind of membership they hold, such as Standard, Junior, Family or Life. A new member starts as Standard.',
        Joined: 'The day they first joined.',
        'Renewal Due': 'The day their next subscription is due. Move it on a year each time they renew.',
        Active: 'Ticked while they are a member. A new member starts ticked; untick it when someone leaves rather than deleting them.',
      } },
      Payments: { about: 'Every payment a member makes: subscriptions, joining fees, event tickets.', fields: {
        Member: 'Who paid, chosen from the list of members by name. It must be filled in.',
        'Paid On': 'The day the money came in.',
        Amount: 'How much they paid.',
        For: 'What it was for, such as 2026 subscription, Joining fee or Annual dinner.',
      } },
    },
    queries: {
      'Renewals due': 'Active members who have a renewal date, the soonest first, with their name, e-mail and membership type. Members who have left (Active unticked) and members with no date are left out. Those at the top whose date has passed are overdue.',
    },
    reports: {
      'Payments by member': 'Each member\'s payments by the day they were paid, with what each was for and its amount, and the total for each member and for everyone. Use it for the treasurer\'s report or to answer a member who asks what they have paid.',
    },
    documents: {},
    tutorial: [
      { title: 'Add a new member', steps: [
        'Open the Members form and press New Record.',
        'Type First Name, Last Name, E-mail, Phone and Address. Change Membership Type if they are not Standard.',
        'Type today in Joined and the day their first renewal falls in Renewal Due, usually a year on. Leave Active ticked and press Add Record.',
      ] },
      { title: 'Record a payment', steps: [
        'Open the Payments form and press New Record.',
        'Choose the member by name in Member, type the day in Paid On, the sum in Amount and what it was for in For.',
        'Press Add Record. The payment also appears under the member on the Members form.',
      ] },
      { title: 'Chase this month\'s renewals', steps: [
        'Run the query Renewals due from the Queries list. The members whose renewal is soonest, or already past, are at the top.',
        'Use the E-mail column to send each one a reminder.',
        'When someone pays, record the payment, then open the Members form at that member and move Renewal Due on a year.',
      ] },
      { title: 'Mark someone who has left', steps: [
        'Open the Members form and move to the member with the Next record arrow.',
        'Untick Active. They drop out of Renewals due but their payments stay on record.',
        'If they come back, tick Active again and type a new date in Renewal Due.',
      ] },
      { title: 'Print the treasurer\'s report', steps: [
        'Open the report Payments by member from the Reports list.',
        'Each member\'s payments are listed under their name with a total, and the grand total is at the foot.',
        'Press Print for the meeting, or Save as PDF to send to the committee.',
      ] },
    ],
    notes: [
      'Member on a payment is a list of names on the Payments form, taken from the Members table by first and last name. Add the member before their first payment. In the Payments table\'s rows the same field shows the member\'s number instead.',
      'Renewal Due is not moved on by itself when a payment comes in. Move it on a year on the member\'s record each time they renew.',
      'A member who has payments cannot be deleted while those payments point at them. Untick Active instead, which also keeps your records of who paid what complete.',
      'Membership Type is typed text. Type each type the same way every time (Junior, not Juniors) so you can choose by it in a query.',
      'Members\' details are personal data. Keep the file somewhere only the committee can open it, and keep a copy with File, Save As.',
    ],
    faq: [
      { q: 'How do I see who has not renewed?', a: 'Run Renewals due. Anyone whose Renewal Due date is before today has not renewed yet, and they are at the top of the list.' },
      { q: 'Can I record a family on one membership?', a: 'Enter one member for the person who pays, with Membership Type Family, and put the other names in a field you add, or enter each person and record the payment against one of them.' },
      { q: 'What if a member pays in instalments?', a: 'Record each instalment as its own payment with the same words in For. The report adds them up under the member.' },
      { q: 'How do I give everyone the same renewal date?', a: 'Type the club\'s renewal day in Renewal Due for each member, such as 1 April. The query lists them all together on that day.' },
      { q: 'Why is a member missing from Renewals due?', a: 'Either Active is unticked or Renewal Due is empty. The query leaves out both.' },
      { q: 'How do I see one member\'s payments quickly?', a: 'Open the Members form at that member: their payments are listed below their details.' },
    ],
    empty: 'Enter your Members first, with the type each holds, the day they joined and when their renewal is due. Then add their Payments, choosing each member from the list. If you are moving from a paper list, start with the current members and add past payments only if you need them.',
    adapt: [
      'Keep the fee for each type: add a Membership Types table with a name and an annual fee, and change Membership Type on Members into a field that points at it.',
      'Record how they paid: add a Method field to Payments (Cash, Card, Bank transfer, Cheque) and a query that totals each method.',
      'Note consent for mailings: add a Mailing Consent field (Yes/No) and the day it was given, and choose by it before you send news.',
      'Track attendance: add an Events table and an Attendance table that points at both Members and Events.',
      'Keep an emergency contact for a sports club: add Emergency Name and Emergency Phone fields to Members.',
      'Add a Left On date to Members alongside Active, to see how long people stay.',
    ],
  },

  crm: {
    name: 'Customers and sales',
    kind: 'Sales tracking',
    line: 'Companies you sell to, the people there, the deals in hand and each call or meeting.',
    summary: [
      'Customers and sales is a simple customer list and sales tracker for a business that sells to other businesses: a consultant, a small agency, a supplier to the trade or a sales team of a few people.',
      'It keeps the companies you deal with and the people at each one, the opportunities you are working on with their stage, value and the day you expect them to close, and a log of every call, e-mail or meeting with a person. A query lists the deals still open by close date, and a report shows the pipeline under each company.',
      'People point at the company they work for, and opportunities point at a company too. Activities point at a person, so a call is kept against whoever you spoke to, and through them against their company.',
    ],
    suits: ['A consultant or freelancer winning work from businesses', 'A small marketing or design agency', 'A supplier to shops, offices or the trade', 'A sales team of two to ten people', 'A training company selling courses to employers', 'A charity seeking corporate sponsors'],
    tables: {
      Companies: { about: 'The businesses you sell to or hope to, one record each.', fields: {
        Name: 'The company\'s name. It must be filled in, and it is what the Company list on a person or opportunity shows.',
        Industry: 'What line of business they are in, such as Retail, Construction or Education.',
        Website: 'Their web address.',
        Phone: 'Their main switchboard number.',
        City: 'The town or city they are based in.',
        Country: 'The country they are based in.',
      } },
      People: { about: 'The people you deal with, each at a company.', fields: {
        Company: 'The company they work for, chosen from your Companies. It can be left empty for someone between jobs.',
        'First Name': 'Their first or given name.',
        'Last Name': 'Their family name. It must be filled in.',
        'Job Title': 'Their job, such as Buyer, Office Manager or Director.',
        'E-mail': 'Their e-mail address.',
        Phone: 'Their direct line or mobile.',
      } },
      Opportunities: { about: 'Each piece of business you are trying to win, from the first enquiry to won or lost.', fields: {
        Company: 'The company the deal is with, chosen from your Companies. It must be filled in.',
        Title: 'A short name for the deal, such as New website or 200 chairs for the hall. It must be filled in.',
        Stage: 'Where it has got to: it starts as Lead, then perhaps Qualified, Proposal or Negotiation, and ends as Won or Lost.',
        Value: 'What the deal is worth if you win it.',
        'Expected Close': 'The day you expect to hear yes or no.',
        Owner: 'Who in your business is looking after it, typed by name.',
      } },
      Activities: { about: 'Every call, e-mail or meeting, kept against the person it was with.', fields: {
        Person: 'Who you spoke to or met, chosen from People by name.',
        When: 'The day and time it happened, or is booked for.',
        Kind: 'What it was: Call (the start), E-mail, Meeting or Visit.',
        Notes: 'What was said, what was agreed and what to do next.',
      } },
    },
    queries: {
      'Open opportunities': 'Deals that are neither Won nor Lost, the soonest expected close first, with the company, the title, the stage, the value and who owns it. Use it each week to see what needs a push.',
    },
    reports: {
      'Pipeline by company': 'Every opportunity under its company, by expected close, with its title, stage and value, and the value added up for each company and for the whole. It includes won and lost deals as well as open ones, so it shows all the business you have had with each company.',
    },
    documents: {},
    tutorial: [
      { title: 'Add a company and the person you deal with', steps: [
        'Open the Companies form, press New Record, type the Name, Industry, Website, Phone, City and Country, and press Add Record.',
        'Open the People form and press New Record. Choose the company in Company, type First Name, Last Name, Job Title, E-mail and Phone, and press Add Record.',
        'Open the Companies form again at the company: the person you added is listed below its fields.',
      ] },
      { title: 'Record a new opportunity', steps: [
        'Open the Opportunities form and press New Record.',
        'Choose the company in Company and type a short Title. Stage starts as Lead.',
        'Type what it is worth in Value, the day you expect an answer in Expected Close and your name in Owner, then press Add Record.',
      ] },
      { title: 'Log a call or meeting', steps: [
        'Open the Activities form and press New Record.',
        'Choose who you spoke to in Person and type the day and time in When.',
        'Change Kind if it was not a call, type what was said and agreed in Notes, and press Add Record.',
      ] },
      { title: 'Move a deal on and close it', steps: [
        'Open the Opportunities form and move to the deal with the Next record arrow.',
        'Change Stage as it progresses, such as from Lead to Proposal, and change Expected Close or Value if they have moved.',
        'When you hear, type Won or Lost in Stage exactly so. It then drops out of Open opportunities.',
      ] },
      { title: 'Review the week\'s pipeline', steps: [
        'Run the query Open opportunities from the Queries list. The deals due to close soonest are at the top.',
        'Open the report Pipeline by company for each company\'s deals with their total value.',
        'Press Print or Save as PDF to take it to a sales meeting.',
      ] },
    ],
    notes: [
      'Company on a person and on an opportunity, and Person on an activity, are lists of names on the form, taken from the other table. Add the company before its people and deals, and the person before their calls. In the table\'s rows these fields show the record\'s number instead.',
      'Stage is typed text. The Open opportunities query leaves out exactly Won and Lost, so a deal typed as won or Closed won stays on it. Keep to the same words each time.',
      'Activities are kept against a person, not a company or a deal. To see every call with a company, open the People form at each of its people: their activities are listed below.',
      'A company that has people or opportunities cannot be deleted while they point at it, and nor can a person with activities.',
      'Keep a copy of the file with File, Save As before you share it with a colleague, so you each know which is the latest.',
    ],
    faq: [
      { q: 'What stages should I use?', a: 'Any that suit how you sell. A common set is Lead, Qualified, Proposal, Negotiation, Won and Lost. Only Won and Lost matter to the query, which leaves them out.' },
      { q: 'How do I see the total value of open deals?', a: 'Make a new report from the query Open opportunities with a total on Value. The report Pipeline by company adds up won and lost deals too, so its totals are larger.' },
      { q: 'Can I link an activity to a deal?', a: 'Not as it comes. Add an Opportunity field to Activities that points at Opportunities, and choose the deal as you log the call.' },
      { q: 'Can more than one person use it?', a: 'It is one file on one computer. To share it, keep it in a shared folder and agree who changes it when, or send a copy with File, Save As.' },
      { q: 'How do I see everyone at one company?', a: 'Open the Companies form at that company: the people and opportunities that point at it are listed below.' },
      { q: 'Why is Owner not a list of names?', a: 'Owner is typed text so the template works for one person as well as a team. Add a Staff table and point Owner at it if you want a list to choose from.' },
    ],
    empty: 'Enter your Companies first, then the People at each, choosing the company from the list. Then add your Opportunities, choosing the company again, and log Activities as you make calls and hold meetings. Start with the deals open now and add closed ones later if you want the history.',
    adapt: [
      'Keep stages to a list: add a Stages table and change Stage into a field that points at it, so it cannot be mistyped.',
      'Weight the pipeline: add a Probability field to Opportunities and a query with a calculated column for Value times Probability.',
      'Link calls to deals: add an Opportunity field to Activities that points at Opportunities.',
      'Plan the next step: add a Follow Up date to Activities and a query of follow-ups in the next 7 days.',
      'Record why deals are lost: add a Lost Reason field to Opportunities and a query that counts each reason.',
    ],
  },

  library: {
    name: 'Lending library',
    kind: 'Lending',
    line: 'Books, the people who borrow them, each loan and what is out and when it is due back.',
    summary: [
      'Lending library is for anyone who lends books and wants them back: a school or classroom library, a church or community book box, a club\'s shared collection, an office reading shelf or a family that lends to friends.',
      'It keeps your books with their author, publisher, year, shelf and number of copies, the people who borrow them, and a record of each loan with the day it went out, the day it is due back and the day it came back. A query lists what is still out by due date, and a report shows each borrower\'s loans.',
      'Each loan points at one book and one borrower. A book can be lent many times, and a borrower can have many loans, so the Loans table is where the two meet.',
    ],
    suits: ['A school or classroom library', 'A church, village hall or community book box', 'A book club sharing its copies', 'An office or staff room reading shelf', 'A toy or tool library, with the fields renamed', 'A family keeping track of books lent to friends'],
    tables: {
      Books: { about: 'One record for each title you lend.', fields: {
        Title: 'The book\'s title. It must be filled in, and it is what the Book list on a loan shows.',
        Author: 'Who wrote it, family name first if you want them to sort that way.',
        ISBN: 'The number on the back cover above the barcode, useful when two books share a title.',
        Publisher: 'Who published it.',
        Year: 'The year it was published.',
        Shelf: 'Where it is kept, such as Fiction A-F or Box 3.',
        Copies: 'How many copies you hold. It starts at 1.',
      } },
      Borrowers: { about: 'The people who borrow from you, one record each.', fields: {
        'First Name': 'Their first or given name.',
        'Last Name': 'Their family name. It must be filled in.',
        'E-mail': 'Where to send a reminder.',
        Phone: 'The number to ring when a book is overdue.',
      } },
      Loans: { about: 'Each time a book goes out, one record per book lent.', fields: {
        Book: 'The book lent, chosen from your Books by title. It must be filled in.',
        Borrower: 'Who borrowed it, chosen from your Borrowers by name. It must be filled in.',
        'Lent On': 'The day it went out.',
        'Due Back': 'The day it should come back.',
        'Returned On': 'The day it came back. Empty while the book is still out.',
      } },
    },
    queries: {
      'On loan': 'Every book still out, which is every loan with no Returned On date, the soonest due back first, with the title and the borrower\'s name and phone. Those at the top whose date has passed are overdue.',
    },
    reports: {
      'Loans by borrower': 'Every loan under its borrower, by the day it was lent, with the book, when it was due and when it came back, and a count of books for each borrower and for the whole. Returned loans are included, so it is each person\'s full borrowing history.',
    },
    documents: {},
    tutorial: [
      { title: 'Add a book', steps: [
        'Open the Books form from the Forms list and press New Record.',
        'Type the Title, Author and ISBN, and the Publisher and Year if you have them.',
        'Type where it lives in Shelf, change Copies if you hold more than one, and press Add Record.',
      ] },
      { title: 'Add a borrower and lend them a book', steps: [
        'Open the Borrowers form, press New Record, type First Name, Last Name, E-mail and Phone, and press Add Record.',
        'Open the Loans form and press New Record. Choose the book in Book and the person in Borrower.',
        'Type today in Lent On and the return day in Due Back, such as two weeks on. Leave Returned On empty and press Add Record.',
      ] },
      { title: 'See what is overdue', steps: [
        'Run the query On loan from the Queries list.',
        'The books due back soonest are at the top; any whose Due Back is before today are overdue.',
        'Use the Phone column to ring the borrower, or look up their e-mail on the Borrowers form.',
      ] },
      { title: 'Record a return', steps: [
        'Open the Loans form and move to the loan with the Next record and Previous record arrows. The Books form at that book also lists its loans below.',
        'Type today in Returned On and press Enter. The book drops out of the On loan query.',
        'Put the book back on the shelf named in its Shelf field on the Books form.',
      ] },
      { title: 'Print the loans report', steps: [
        'Open the report Loans by borrower from the Reports list.',
        'Each borrower\'s loans are listed under their name with a count of books.',
        'Press Print, or Save as PDF to keep a copy for the end of term.',
      ] },
    ],
    notes: [
      'Book and Borrower on a loan are lists of names on the Loans form, taken from Books by title and from Borrowers by first and last name. Add the book and the borrower before the loan. In the Loans table\'s rows these fields show the record\'s number instead.',
      'A loan is never deleted when the book comes back: typing Returned On is what closes it. That keeps the history the report prints.',
      'Copies is a count you keep for your own reference. The template does not stop you lending more copies than you hold, so check On loan before lending the last one.',
      'A book or borrower that has loans cannot be deleted while those loans point at it.',
      'Keep a copy of the file with File, Save As, at the end of each term or season.',
    ],
    faq: [
      { q: 'How do I see who has a particular book?', a: 'Open the Books form at that book: its loans are listed below, and the one with no Returned On is out now.' },
      { q: 'How do I lend two books at once?', a: 'Add a loan for each book, both with the same borrower and dates.' },
      { q: 'Can I renew a loan?', a: 'Yes. Open the loan on the Loans form and type the new date in Due Back.' },
      { q: 'Two books have the same title. How do I tell them apart in the list?', a: 'The Book list shows titles only. Add the author or year to the Title, such as Collected Poems (Hardy), or keep each by its ISBN.' },
      { q: 'Can I lend things other than books?', a: 'Yes. Rename the fields in the Books table\'s design, such as Author to Make, and use it for toys, tools or games.' },
      { q: 'Why is a book still showing as on loan?', a: 'Its loan has no Returned On date. Open the loan and type the day it came back.' },
    ],
    empty: 'Enter your Books first, then your Borrowers, then add a Loan each time a book goes out, choosing the book and the borrower from the lists. If books are out already, add a loan for each with the day it went out and the day it is due.',
    adapt: [
      'Set a loan length for each book: add a Loan Days field to Books (14 for most, 7 for books in demand) and use it when you type Due Back.',
      'Charge late fees: add a Fine field to Loans and a query of returned loans whose Returned On is later than Due Back.',
      'Group the borrowers: add a Class or Group field to Borrowers so a school can print each class\'s loans.',
      'Sort by subject: add a Subject or Genre field to Books and a query that lists the books in one subject.',
      'Take reservations: add a Reservations table that points at Books and Borrowers, with the day it was asked for.',
    ],
  },
};
