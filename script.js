fetch("data/superstore.xlsx")
  .then((response) => response.arrayBuffer())

  .then((data) => {
    const workbook = XLSX.read(data, {
      cellDates: true,
    });

    const sheet = workbook.Sheets["Orders"];

    const jsonData = XLSX.utils.sheet_to_json(sheet);

    // KPI

    let totalSales = 0;
    let totalProfit = 0;
    let totalDiscount = 0;

    let orders = new Set();

    // SALES PROFIT

    const yearlySales = {};
    const yearlyProfit = {};

    // CATEGORY

    const categoryProfit = {};

    // DISCOUNT

    const discountImpact = {
      Low: 0,
      Medium: 0,
      High: 0,
    };

    // LOOP

    jsonData.forEach((row) => {
      const sales = Number(row["Sales"]) || 0;

      const profit = Number(row["Profit"]) || 0;

      const discount = Number(row["Discount"]) || 0;

      totalSales += sales;
      totalProfit += profit;
      totalDiscount += discount;

      // ORDER

      if (row["Order ID"]) {
        orders.add(row["Order ID"]);
      }

      // DATE

      const date = row["Order Date"];

      if (date instanceof Date) {
        const year = date.getFullYear();

        if (!yearlySales[year]) {
          yearlySales[year] = 0;
        }

        if (!yearlyProfit[year]) {
          yearlyProfit[year] = 0;
        }

        yearlySales[year] += sales;
        yearlyProfit[year] += profit;
      }

      // CATEGORY

      const category = row["Category"];

      if (!categoryProfit[category]) {
        categoryProfit[category] = 0;
      }

      categoryProfit[category] += profit;

      // DISCOUNT

      if (discount <= 0.1) {
        discountImpact["Low"] += profit;
      } else if (discount <= 0.3) {
        discountImpact["Medium"] += profit;
      } else {
        discountImpact["High"] += profit;
      }
    });

    // KPI DISPLAY

    document.getElementById("totalSales").innerHTML =
      "$" + (totalSales / 1000000).toFixed(2) + "M";

    document.getElementById("totalProfit").innerHTML =
      "$" + Math.round(totalProfit / 1000) + "K";

    document.getElementById("totalOrders").innerHTML =
      orders.size.toLocaleString();

    document.getElementById("avgDiscount").innerHTML =
      ((totalDiscount / jsonData.length) * 100).toFixed(1) + "%";

    // CHART 1

    const salesProfitChart = document.getElementById("salesProfitChart");

    new Chart(salesProfitChart, {
      type: "line",

      data: {
        labels: Object.keys(yearlySales),

        datasets: [
          {
            label: "Sales",
            data: Object.values(yearlySales),
            borderColor: "#2563eb",
            backgroundColor: "#2563eb",
            tension: 0.4,
          },

          {
            label: "Profit",
            data: Object.values(yearlyProfit),
            borderColor: "#16a34a",
            backgroundColor: "#16a34a",
            tension: 0.4,
          },
        ],
      },

      options: {
        responsive: true,
      },
    });

    // CHART 2

    const categoryChart = document.getElementById("categoryChart");

    new Chart(categoryChart, {
      type: "bar",

      data: {
        labels: Object.keys(categoryProfit),

        datasets: [
          {
            label: "Profit",

            data: Object.values(categoryProfit),

            backgroundColor: ["#ef4444", "#64748b", "#16a34a"],
          },
        ],
      },

      options: {
        responsive: true,
      },
    });

    // CHART 3

    const discountChart = document.getElementById("discountChart");

    new Chart(discountChart, {
      type: "doughnut",

      data: {
        labels: Object.keys(discountImpact),

        datasets: [
          {
            data: Object.values(discountImpact),

            backgroundColor: ["#16a34a", "#facc15", "#ef4444"],
          },
        ],
      },

      options: {
        responsive: true,
      },
    });
  });
