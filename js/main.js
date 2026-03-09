/**
 * Fuel Usage Per Day — MyGeotab Add-In
 * Database: sasa_inti
 *
 * Calculates daily fuel usage for all vehicles using the FuelUsed API object.
 */
geotab.addin.fuelUsagePerDay = function () {
    "use strict";

    var api, state;
    var deviceCache = {};

    // ---- DOM helpers ----
    function $(id) { return document.getElementById(id); }

    function showLoading(show) {
        $("loading").style.display = show ? "block" : "none";
    }

    function showError(msg) {
        var el = $("errorMsg");
        if (msg) {
            el.textContent = msg;
            el.style.display = "block";
        } else {
            el.style.display = "none";
        }
    }

    // ---- Set default dates (last 7 days) ----
    function setDefaultDates() {
        var today = new Date();
        var weekAgo = new Date();
        weekAgo.setDate(today.getDate() - 7);

        $("toDate").value = formatDateInput(today);
        $("fromDate").value = formatDateInput(weekAgo);
    }

    function formatDateInput(d) {
        var yyyy = d.getFullYear();
        var mm = String(d.getMonth() + 1).padStart(2, "0");
        var dd = String(d.getDate()).padStart(2, "0");
        return yyyy + "-" + mm + "-" + dd;
    }

    function formatDateDisplay(dateStr) {
        var d = new Date(dateStr);
        return d.getFullYear() + "-" +
            String(d.getMonth() + 1).padStart(2, "0") + "-" +
            String(d.getDate()).padStart(2, "0");
    }

    // ---- Load all devices into cache ----
    function loadDevices(callback) {
        api.call("Get", {
            typeName: "Device"
        }, function (devices) {
            deviceCache = {};
            devices.forEach(function (d) {
                deviceCache[d.id] = d.name || d.serialNumber || d.id;
            });
            callback(devices);
        }, function (err) {
            showError("Failed to load vehicles: " + err.message);
            showLoading(false);
        });
    }

    // ---- Load fuel data ----
    function loadFuelData() {
        var fromDate = $("fromDate").value;
        var toDate = $("toDate").value;

        if (!fromDate || !toDate) {
            showError("Please select both From and To dates.");
            return;
        }

        if (new Date(fromDate) > new Date(toDate)) {
            showError("'From' date must be before 'To' date.");
            return;
        }

        showError(null);
        showLoading(true);
        $("tableWrapper").style.display = "none";
        $("summaryCards").style.display = "none";

        // Load devices first, then fuel data
        loadDevices(function () {
            api.call("Get", {
                typeName: "FuelUsed",
                search: {
                    fromDate: fromDate + "T00:00:00.000Z",
                    toDate: toDate + "T23:59:59.999Z"
                }
            }, function (fuelRecords) {
                showLoading(false);
                processFuelData(fuelRecords);
            }, function (err) {
                showLoading(false);
                showError("Failed to load fuel data: " + err.message);
            });
        });
    }

    // ---- Process & aggregate fuel data by vehicle + day ----
    function processFuelData(records) {
        if (!records || records.length === 0) {
            showError("No fuel data found for the selected date range.");
            return;
        }

        // Group by deviceId + date
        var grouped = {};
        records.forEach(function (r) {
            var deviceId = r.device ? r.device.id : "Unknown";
            var dateKey = formatDateDisplay(r.dateTime);
            var key = deviceId + "|" + dateKey;

            if (!grouped[key]) {
                grouped[key] = {
                    deviceId: deviceId,
                    vehicleName: deviceCache[deviceId] || deviceId,
                    date: dateKey,
                    totalFuel: 0,
                    idleFuel: 0
                };
            }

            grouped[key].totalFuel += (r.totalFuelUsed || 0);
            grouped[key].idleFuel += (r.totalIdlingFuelUsedL || 0);
        });

        // Convert to sorted array
        var rows = Object.values(grouped);
        rows.sort(function (a, b) {
            if (a.date === b.date) return a.vehicleName.localeCompare(b.vehicleName);
            return a.date.localeCompare(b.date);
        });

        // Calculate summary
        var uniqueVehicles = new Set(rows.map(function (r) { return r.deviceId; }));
        var totalFuelAll = rows.reduce(function (sum, r) { return sum + r.totalFuel; }, 0);
        var totalIdleAll = rows.reduce(function (sum, r) { return sum + r.idleFuel; }, 0);
        var uniqueDays = new Set(rows.map(function (r) { return r.date; }));
        var avgPerVehiclePerDay = uniqueVehicles.size > 0 && uniqueDays.size > 0
            ? totalFuelAll / uniqueVehicles.size / uniqueDays.size
            : 0;

        // Update summary cards
        $("totalVehicles").textContent = uniqueVehicles.size;
        $("totalFuel").textContent = totalFuelAll.toFixed(2);
        $("avgFuel").textContent = avgPerVehiclePerDay.toFixed(2);
        $("totalIdleFuel").textContent = totalIdleAll.toFixed(2);
        $("summaryCards").style.display = "flex";

        // Render table
        renderTable(rows);
    }

    // ---- Render table ----
    function renderTable(rows) {
        var tbody = $("fuelTableBody");
        tbody.innerHTML = "";

        rows.forEach(function (row) {
            var drivingFuel = row.totalFuel - row.idleFuel;
            var tr = document.createElement("tr");
            tr.innerHTML =
                "<td>" + escapeHtml(row.vehicleName) + "</td>" +
                "<td>" + row.date + "</td>" +
                "<td>" + row.totalFuel.toFixed(2) + "</td>" +
                "<td>" + row.idleFuel.toFixed(2) + "</td>" +
                "<td>" + drivingFuel.toFixed(2) + "</td>";
            tbody.appendChild(tr);
        });

        $("tableWrapper").style.display = "block";
    }

    function escapeHtml(str) {
        var div = document.createElement("div");
        div.textContent = str;
        return div.innerHTML;
    }

    // ---- Export CSV ----
    function exportCsv() {
        var table = $("fuelTable");
        if (!table) return;

        var rows = table.querySelectorAll("tr");
        var csvLines = [];

        rows.forEach(function (row) {
            var cols = row.querySelectorAll("th, td");
            var line = [];
            cols.forEach(function (col) {
                var text = col.textContent.replace(/"/g, '""');
                line.push('"' + text + '"');
            });
            csvLines.push(line.join(","));
        });

        var csvContent = csvLines.join("\n");
        var blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
        var url = URL.createObjectURL(blob);
        var link = document.createElement("a");
        link.href = url;
        link.download = "fuel_usage_per_day_sasa_inti.csv";
        link.click();
        URL.revokeObjectURL(url);
    }

    // ---- Add-In Lifecycle ----
    return {
        initialize: function (freshApi, freshState, callback) {
            api = freshApi;
            state = freshState;

            // Set default dates
            setDefaultDates();

            // Bind button events
            $("btnLoad").addEventListener("click", loadFuelData);
            $("btnExport").addEventListener("click", exportCsv);

            callback();
        },

        focus: function (freshApi, freshState) {
            api = freshApi;
            state = freshState;

            // Auto-load data on focus
            loadFuelData();
        },

        blur: function () {
            // Clean up if needed
        }
    };
};
