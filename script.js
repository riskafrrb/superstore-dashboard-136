fetch("data/superstore.xlsx")
  .then((response) => response.arrayBuffer())

  .then((data) => {
    // FIX EXCEL DATE
    const workbook = XLSX.read(data, {
      cellDates: true,
    });

    // AMBIL SHEET ORDERS
    const sheet = workbook.Sheets["Orders"];

    // CONVERT JSON
    const jsonData = XLSX.utils.sheet_to_json(sheet);

    console.log(jsonData[0]);

    // KPI
    let totalSales = 0;
    let totalProfit = 0;

    let orders = new Set();
    let customers = new Set();

    // CARI TAHUN TERBARU
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

    // REVENUE BULANAN
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

        // FILTER TAHUN TERBARU
        if (year === latestYear) {
          const month = date.toLocaleString("default", {
            month: "short",
          });

          if (monthlyRevenue[month] !== undefined) {
            monthlyRevenue[month] += sales;
          }
        }
      }
    });

    // KPI DISPLAY

    document.getElementById("totalSales").innerHTML =
      "$" + Math.round(totalSales).toLocaleString();

    document.getElementById("totalOrders").innerHTML = orders.size;

    document.getElementById("totalProfit").innerHTML =
      "$" + Math.round(totalProfit).toLocaleString();

    document.getElementById("totalCustomers").innerHTML = customers.size;

    // TITLE DINAMIS

    document.getElementById("revenueTitle").innerHTML =
      `Revenue by Month -- ${latestYear}`;

    // REVENUE BARS

    const revenueBars = document.getElementById("revenueBars");

    const maxRevenue = Math.max(...Object.values(monthlyRevenue));

    const monthNames = {
      Jan: `Januari ${latestYear}`,
      Feb: `Februari ${latestYear}`,
      Mar: `Maret ${latestYear}`,
      Apr: `April ${latestYear}`,
      May: `Mei ${latestYear}`,
      Jun: `Juni ${latestYear}`,
      Jul: `Juli ${latestYear}`,
      Aug: `Agustus ${latestYear}`,
      Sep: `September ${latestYear}`,
      Oct: `Oktober ${latestYear}`,
      Nov: `November ${latestYear}`,
      Dec: `Desember ${latestYear}`,
    };

    Object.keys(monthlyRevenue).forEach((month) => {
      const revenue = monthlyRevenue[month];

      const percentage = maxRevenue > 0 ? (revenue / maxRevenue) * 100 : 0;

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
