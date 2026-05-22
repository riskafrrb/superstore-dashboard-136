fetch("data/superstore.xlsx")
  .then((response) => response.arrayBuffer())

  .then((data) => {
    const workbook = XLSX.read(data, {
      cellDates: true,
    });

    const sheet = workbook.Sheets["Orders"];

    const jsonData = XLSX.utils.sheet_to_json(sheet);

    console.log(jsonData[0]);

    let totalSales = 0;
    let totalProfit = 0;

    let orders = new Set();
    let customers = new Set();
    let latestYear = 0;

    jsonData.forEach((row) => {
      const date = row["Order Date"];

      if (date instanceof Date) {
        const year = date.getFullYear();

        if (year > latestYear) {
          latestYear = year;
        }
      }
    });
    const monthlyRevenue = {
      Jan: 0,
      Feb: 0,
      Mar: 0,
      Apr: 0,
      May: 0,
      Jun: 0,
      Jul: 0,
      Aug: 0,
      Sep: 0,
      Oct: 0,
      Nov: 0,
      Dec: 0,
    };

    jsonData.forEach((row) => {
      const sales = Number(row["Sales"]) || 0;
      const profit = Number(row["Profit"]) || 0;

      totalSales += sales;
      totalProfit += profit;

      // ORDER
      if (row["Order ID"]) {
        orders.add(row["Order ID"]);
      }

      // CUSTOMER
      if (row["Customer ID"]) {
        customers.add(row["Customer ID"]);
      }

      // DATE
      const date = row["Order Date"];

      if (date instanceof Date) {
        const year = date.getFullYear();

        if (year === 2024) {
          const month = date.toLocaleString("default", {
            month: "short",
          });

          if (monthlyRevenue[month] !== undefined) {
            monthlyRevenue[month] += sales;
          }
        }
      }
    });

    document.getElementById("totalSales").innerHTML =
      "$" + Math.round(totalSales).toLocaleString();

    document.getElementById("totalOrders").innerHTML = orders.size;

    document.getElementById("totalProfit").innerHTML =
      "$" + Math.round(totalProfit).toLocaleString();

    document.getElementById("totalCustomers").innerHTML = customers.size;

    const revenueBars = document.getElementById("revenueBars");

    const maxRevenue = Math.max(...Object.values(monthlyRevenue));

    const monthNames = {
      Jan: "Januari 2024",
      Feb: "Februari 2024",
      Mar: "Maret 2024",
      Apr: "April 2024",
      May: "Mei 2024",
      Jun: "Juni 2024",
      Jul: "Juli 2024",
      Aug: "Agustus 2024",
      Sep: "September 2024",
      Oct: "Oktober 2024",
      Nov: "November 2024",
      Dec: "Desember 2024",
    };

    Object.keys(monthlyRevenue).forEach((month) => {
      const revenue = monthlyRevenue[month];

      const percentage = (revenue / maxRevenue) * 100;

      revenueBars.innerHTML += `

        <div class="revenue-item">

            <div class="month-label">
                ${monthNames[month]}
            </div>

            <div class="bar-container">

                <div class="bar"
                    style="width:${percentage}%">

                    $${Math.round(revenue).toLocaleString()}

                </div>

            </div>

            <div class="amount">
                $${Math.round(revenue / 1000)}K
            </div>

        </div>

        `;
    });
  });
