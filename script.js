fetch("data/superstore.xlsx")
  .then((response) => response.arrayBuffer())

  .then((data) => {
    const workbook = XLSX.read(data, {
      cellDates: true,
    });

    const sheet = workbook.Sheets["Orders"];

    const jsonData = XLSX.utils.sheet_to_json(sheet);

    let totalSales = 0;
    let totalProfit = 0;

    let customers = new Set();

    const regionSales = {};
    const regionProfit = {};
    const regionCustomers = {};
    const regionMargin = {};
    const yearlyRegionSales = {};

    jsonData.forEach((row) => {
      const region = row["Region"];

      const sales = Number(row["Sales"]) || 0;

      const profit = Number(row["Profit"]) || 0;

      totalSales += sales;

      totalProfit += profit;

      customers.add(row["Customer ID"]);

      if (!regionSales[region]) {
        regionSales[region] = 0;
      }

      regionSales[region] += sales;

      if (!regionProfit[region]) {
        regionProfit[region] = 0;
      }

      regionProfit[region] += profit;

      if (!regionCustomers[region]) {
        regionCustomers[region] = new Set();
      }

      regionCustomers[region].add(row["Customer ID"]);

      const date = row["Order Date"];

      if (date instanceof Date) {
        const year = date.getFullYear();

        if (!yearlyRegionSales[year]) {
          yearlyRegionSales[year] = {};
        }

        if (!yearlyRegionSales[year][region]) {
          yearlyRegionSales[year][region] = 0;
        }

        yearlyRegionSales[year][region] += sales;
      }
    });

    Object.keys(regionSales).forEach((region) => {
      regionMargin[region] = (regionProfit[region] / regionSales[region]) * 100;
    });

    document.getElementById("totalSales").innerHTML =
      "$" + (totalSales / 1000000).toFixed(2) + "M";

    document.getElementById("totalProfit").innerHTML =
      "$" + (totalProfit / 1000).toFixed(0) + "K";

    document.getElementById("totalCustomers").innerHTML = customers.size;

    document.getElementById("avgMargin").innerHTML =
      ((totalProfit / totalSales) * 100).toFixed(1) + "%";

    const tooltipConfig = {
      callbacks: {
        label: function (context) {
          return "$" + context.raw.toLocaleString();
        },
      },
    };

    new Chart(document.getElementById("revenueRegionChart"), {
      type: "bar",

      data: {
        labels: Object.keys(regionSales),

        datasets: [
          {
            label: "Revenue",

            data: Object.values(regionSales),

            backgroundColor: Object.keys(regionSales).map((region) =>
              region === "West" ? "#7c3aed" : "#d6d3d1",
            ),
            borderRadius: 10,
          },
        ],
      },

      options: {
        responsive: true,

        plugins: {
          legend: {
            display: false,
          },

          tooltip: tooltipConfig,
        },
      },
    });

    new Chart(document.getElementById("profitRegionChart"), {
      type: "bar",

      data: {
        labels: Object.keys(regionProfit),

        datasets: [
          {
            label: "Profit",

            data: Object.values(regionProfit),

            backgroundColor: Object.keys(regionProfit).map((region) =>
              region === "West" ? "#059669" : "#d6d3d1",
            ),
            borderRadius: 10,
          },
        ],
      },

      options: {
        responsive: true,

        plugins: {
          legend: {
            display: false,
          },

          tooltip: tooltipConfig,
        },
      },
    });

    new Chart(document.getElementById("marginChart"), {
      type: "bar",

      data: {
        labels: Object.keys(regionMargin),

        datasets: [
          {
            label: "Margin %",

            data: Object.values(regionMargin),

            backgroundColor: Object.keys(regionMargin).map((region) =>
              region === "Central"
                ? "#e11d48"
                : region === "West"
                  ? "#7c3aed"
                  : "#d6d3d1",
            ),
            borderRadius: 10,
          },
        ],
      },

      options: {
        responsive: true,

        plugins: {
          legend: {
            display: false,
          },

          tooltip: {
            callbacks: {
              label: function (context) {
                return context.raw.toFixed(1) + "%";
              },
            },
          },
        },
      },
    });

    const customerLabels = Object.keys(regionCustomers);

    const customerValues = customerLabels.map(
      (region) => regionCustomers[region].size,
    );

    new Chart(document.getElementById("customerChart"), {
      type: "bar",

      data: {
        labels: customerLabels,

        datasets: [
          {
            label: "Customers",

            data: customerValues,

            backgroundColor: "#7c3aed",

            borderRadius: 10,
          },
        ],
      },

      options: {
        indexAxis: "y",

        responsive: true,

        plugins: {
          legend: {
            display: false,
          },
        },
      },
    });

    const years = Object.keys(yearlyRegionSales);

    const westSales = years.map((year) => yearlyRegionSales[year]["West"] || 0);

    new Chart(document.getElementById("growthChart"), {
      type: "line",

      data: {
        labels: years,

        datasets: [
          {
            label: "West Revenue Growth",

            data: westSales,

            borderColor: "#7c3aed",

            backgroundColor: "#7c3aed",

            tension: 0.4,

            borderWidth: 3,
          },
        ],
      },

      options: {
        responsive: true,

        plugins: {
          tooltip: tooltipConfig,
        },
      },
    });
  });
