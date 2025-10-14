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



// 获取页面元素
const editableDiv = document.getElementById("editableText");
const editBtn = document.getElementById("editBtn");
const saveBtn = document.getElementById("saveBtn");
const cancelBtn = document.getElementById("cancelBtn");
let originalText = editableDiv.innerHTML;

// 获取pageId的辅助函数
function getPageId() {
    return document.getElementById('pagePanel').getAttribute('data-page-id');
}

function enableEdit() {
    originalText = editableDiv.innerHTML;
    editableDiv.contentEditable = "true";
    editableDiv.focus();
    editBtn.style.display = "none";
    saveBtn.style.display = "inline-block";
    cancelBtn.style.display = "inline-block";
}

async function saveText() {
    const text = editableDiv.innerText;
    const pageId = getPageId();
    
    try {
        const res = await fetch(`/save/${pageId}`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ text })
        });
        
        const data = await res.json();
        
        if (res.ok) {
            alert(data.message || "保存成功！");
            // 保存成功后更新originalText
            originalText = editableDiv.innerHTML;
            editableDiv.contentEditable = "false";
            saveBtn.style.display = "none";
            cancelBtn.style.display = "none";
            editBtn.style.display = "inline-block";
        } else {
            alert(data.message || "保存失败！");
        }
    } catch (error) {
        console.error('保存错误:', error);
        alert("保存失败，请检查网络连接！");
    }
}

function cancelEdit() {
    editableDiv.innerHTML = originalText;
    editableDiv.contentEditable = "false";
    saveBtn.style.display = "none";
    cancelBtn.style.display = "none";
    editBtn.style.display = "inline-block";
}

function exportWord() {
    const pageId = getPageId();
    window.location.href = `/export/${pageId}`;
}

// 添加键盘事件支持
editableDiv.addEventListener('keydown', function(event) {
    if (event.key === 'Enter' && !event.shiftKey) {
        event.preventDefault();
        // 可选：按Enter键自动保存
        // saveText();
    }
});

document.head.appendChild(style);

// 页面加载完成后的初始化
document.addEventListener('DOMContentLoaded', function() {
    console.log('页面加载完成，当前页面ID:', getPageId());
    
    // 确保editableDiv初始不可编辑
    editableDiv.contentEditable = "false";
    
    // 初始化按钮状态
    editBtn.style.display = "inline-block";
    saveBtn.style.display = "none";
    cancelBtn.style.display = "none";
});