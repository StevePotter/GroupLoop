<!-- Main menus. -->

<div id="glMenu" class="menu"
     onmouseover="menuMouseover(event)">
</div>


<div id="customerMenu" class="menu"
     onmouseover="menuMouseover(event)">
<a class="menuItem" href="customers.asp">Customer List</a>
<div class="menuItemSep"></div>
<a class="menuItem" href="mailing.asp">Mail Customers</a>
</div>


<div id="maintMenu" class="menu"
     onmouseover="menuMouseover(event)">
<a class="menuItem" href="daily_setup.asp">Run the daily maintenance</a>
<a class="menuItem" href="maintenance_modify.asp">Maintenance run list</a>
<div class="menuItemSep"></div>
<a class="menuItem" href="template_copy.asp">Distribute a template file</a>
<div class="menuItemSep"></div>
<a class="menuItem" href="upload.asp">Upload files to the server</a>
</div>


<div id="financialMenu" class="menu">
<a class="menuItem" href=""
   onclick="return false;"
   onmouseover="menuItemMouseover(event, 'invoicesMenu');"
><span class="menuItemText">Invoices</span><span class="menuItemArrow">&#9654;</span></a>
<a class="menuItem" href=""
   onclick="return false;"
   onmouseover="menuItemMouseover(event, 'monthlychargesMenu');"
><span class="menuItemText">Monthly Charges</span><span class="menuItemArrow">&#9654;</span></a>
<div class="menuItemSep"></div>

<a class="menuItem" href=""
   onclick="return false;"
   onmouseover="menuItemMouseover(event, 'depositsMenu');"
><span class="menuItemText">Deposits</span><span class="menuItemArrow">&#9654;</span></a>
<a class="menuItem" href=""
   onclick="return false;"
   onmouseover="menuItemMouseover(event, 'withdrawalsMenu');"
><span class="menuItemText">Withdrawals</span><span class="menuItemArrow">&#9654;</span></a>

<div class="menuItemSep"></div>
<a class="menuItem" href=""
   onclick="return false;"
   onmouseover="menuItemMouseover(event, 'acctsMenu');"
><span class="menuItemText">Bank Accounts</span><span class="menuItemArrow">&#9654;</span></a>

</div>


<!-- financial sub menus. -->

<div id="invoicesMenu" class="menu" onmouseover="menuMouseover(event)">
<a class="menuItem" href="invoice_add.asp">New Invoice</a>
<a class="menuItem" href="invoices_modify.asp">Modify Invoices</a>
</div>

<div id="monthlychargesMenu" class="menu" onmouseover="menuMouseover(event)">
<a class="menuItem" href="monthlycharge_add.asp">New Charge</a>
<a class="menuItem" href="monthlycharges_modify.asp">Modify Charges</a>
</div>

<div id="depositsMenu" class="menu" onmouseover="menuMouseover(event)">
<a class="menuItem" href="bankdeposits_add.asp">Record Deposits</a>
<a class="menuItem" href="bankdeposits_modify.asp">Modify Deposits</a>
</div>

<div id="withdrawalsMenu" class="menu" onmouseover="menuMouseover(event)">
<a class="menuItem" href="bankwithdrawals_add.asp">Record Withdrawals</a>
<a class="menuItem" href="bankwithdrawals_modify.asp">Modify Withdrawals</a>
</div>

<div id="acctsMenu" class="menu" onmouseover="menuMouseover(event)">
<a class="menuItem" href="" onclick="return false;" onmouseover="menuItemMouseover(event, 'bankstatementsMenu');">
<span class="menuItemText">Bank Statements</span><span class="menuItemArrow">&#9654;</span>
</a>
<div class="menuItemSep"></div>
<a class="menuItem" href="" onclick="return false;" onmouseover="menuItemMouseover(event, 'bankaccountsMenu');">
<span class="menuItemText">Modify Bank Accounts</span><span class="menuItemArrow">&#9654;</span>
</a>
</div>

<!-- bank accounts sub menus. -->
<div id="bankstatementsMenu" class="menu" onmouseover="menuMouseover(event)">
<a class="menuItem" href="bankstatements_add.asp">Record New Statements</a>
<a class="menuItem" href="bankstatements_modify.asp">Past Statements</a>
</div>

<div id="bankaccountsMenu" class="menu" onmouseover="menuMouseover(event)">
<a class="menuItem" href="bankaccounts_add.asp">New Account</a>
<a class="menuItem" href="bankaccounts_modify.asp">Modify Accounts</a>
</div>


<!-- employees menu. -->
<div id="employeesMenu" class="menu" onmouseover="menuMouseover(event)">
<a class="menuItem" href="employees_add.asp">New Employee</a>
<a class="menuItem" href="employees_modify.asp">Modify Employees</a>
</div>

<!-- stats menu. -->
<div id="statsMenu" class="menu" onmouseover="menuMouseover(event)">
<a class="menuItem"  href="stats_homesite.asp">GroupLoop.com Home Site</a>
<a class="menuItem"  href="stats_customers.asp">Customers</a>
<a class="menuItem"  href="stats_billing.asp">Financial</a>
</div>