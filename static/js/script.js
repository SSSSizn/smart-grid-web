document.addEventListener("DOMContentLoaded", function () {
    const loading = document.getElementById("loading");
    const tableContent = document.getElementById("tableContent");
    const tableBody = document.getElementById("tableBody");
    const errorMessage = document.getElementById("errorMessage");

    const chartContainer = document.getElementById("chartContainer");
    const chartCanvas = document.getElementById("chartCanvas");
    let chartInstance = null;

    // 获取数据并渲染表格
    fetch("/api/problem-lines")
        .then(res => res.json())
        .then(res => {
            if (!res.success) throw new Error("数据加载失败");
            const data = res.data;

            if (!data.length) {
                errorMessage.style.display = "block";
                errorMessage.innerText = "没有可显示的数据";
                loading.style.display = "none";
                return;
            }

            data.forEach(row => {
                const tr = document.createElement("tr");

                // 根据负载率设置颜色
                const loadRate = parseFloat(row["2024年线路最大负载率（%）"]);
                if (loadRate >= 90) tr.style.backgroundColor = "#f8d7da"; // 红色重载
                else if (loadRate <= 25) tr.style.backgroundColor = "#d4edda"; // 绿色轻载

                for (const key of ["序号","变电站名称","线路名称","限额电流（A）",
                                    "2024年最大电流（A）","2024年线路最大负载率（%）","重载/轻载原因"]) {
                    const td = document.createElement("td");
                    td.innerText = row[key];
                    tr.appendChild(td);
                }
                tableBody.appendChild(tr);
            });

            loading.style.display = "none";
            tableContent.style.display = "block";
        })
        .catch(err => {
            loading.style.display = "none";
            errorMessage.style.display = "block";
            errorMessage.innerText = err;
        });

// 显示柱状图
document.getElementById("showBarChart").addEventListener("click", () => {
    fetch("/api/problem-lines").then(res => res.json()).then(res => {
        const data = res.data;
        const top10 = data.sort((a,b)=>b["2024年线路最大负载率（%）"]-a["2024年线路最大负载率（%）"]).slice(0,10);
        const labels = top10.map(r => r["线路名称"]);
        const values = top10.map(r => r["2024年线路最大负载率（%）"]);

        if(chartInstance) chartInstance.destroy();
        chartInstance = new Chart(chartCanvas, {
            type: "bar",
            data: { labels, datasets: [{ label: "负载率(%)", data: values, backgroundColor:"#0d6efd" }] },
            options: { responsive:true }
        });

        chartContainer.style.display = "block";
        document.getElementById("hideCharts").style.display = "inline-block";

        // 设置图表标题，统一 Word 标题
        document.getElementById("chartTitle").innerText = "前10条线路负载率情况";
    });
});

// 显示饼图
document.getElementById("showPieChart").addEventListener("click", () => {
    fetch("/api/problem-lines").then(res => res.json()).then(res => {
        const data = res.data;
        const heavy = data.filter(r=>r["2024年线路最大负载率（%）"]>=90).length;
        const light = data.filter(r=>r["2024年线路最大负载率（%）"]<=25).length;

        if(chartInstance) chartInstance.destroy();
        chartInstance = new Chart(chartCanvas, {
            type: "pie",
            data: {
                labels:["重载","轻载"],
                datasets:[{ data:[heavy, light], backgroundColor:["#dc3545","#198754"] }]
            },
            options: { responsive:true }
        });

        chartContainer.style.display = "block";
        document.getElementById("hideCharts").style.display = "inline-block";

        // 设置图表标题，统一 Word 标题
        document.getElementById("chartTitle").innerText = "重载/轻载占比";
    });
});
    // 隐藏图表
    document.getElementById("hideCharts").addEventListener("click", () => {
        if(chartInstance) chartInstance.destroy();
        chartContainer.style.display = "none";
        document.getElementById("hideCharts").style.display = "none";
    });
});

