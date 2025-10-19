
// 点击展开/收起子菜单
document.querySelectorAll(".has-submenu").forEach(item => {
  item.addEventListener("click", function (e) {
    e.stopPropagation();
    this.classList.toggle("active");
  });
});

// 点击具体“表”时，加载对应iframe + 显示分析面板
document.querySelectorAll("li[data-url]").forEach(item => {
  item.addEventListener("click", function (e) {
    e.stopPropagation();

    const iframe = document.getElementById("content-frame");
    const placeholder = document.getElementById("placeholder");
    const allPanels = document.querySelectorAll(".analysis-panel");

    // 隐藏所有分析区
    allPanels.forEach(p => p.style.display = "none");

    // 设置iframe内容
    const url = this.getAttribute("data-url");
    iframe.src = url;
    iframe.style.display = "block";
    placeholder.style.display = "none";

    // 显示对应分析区
    const panelId = this.getAttribute("data-panel");
    const panel = document.getElementById(panelId);
    if (panel) panel.style.display = "block";
  });
});

// 在文件末尾添加表格加载处理
document.querySelectorAll("li[data-url]").forEach(item => {
  item.addEventListener("click", function (e) {
    e.stopPropagation();

    const iframe = document.getElementById("content-frame");
    const placeholder = document.getElementById("placeholder");
    const allPanels = document.querySelectorAll(".analysis-panel");

    // 隐藏所有分析区
    allPanels.forEach(p => p.style.display = "none");

    // 设置iframe内容
    const url = this.getAttribute("data-url");
    iframe.src = url;
    iframe.style.display = "block";
    placeholder.style.display = "none";

    // 显示对应分析区
    const panelId = this.getAttribute("data-panel");
    const panel = document.getElementById(panelId);
    if (panel) panel.style.display = "block";
  });
});
// 在script2.js中添加以下代码
document.getElementById("content-frame").onload = function() {
    const iframeDoc = this.contentDocument || this.contentWindow.document;
    const table = iframeDoc.querySelector('table');
    if (!table) return;

    // 获取所有表头
    const headers = Array.from(table.querySelectorAll('thead th'));
    const panel = document.querySelector('.analysis-panel:not([style*="display: none"])');

    if (panel) {
        const buttons = panel.querySelectorAll('button');
        buttons.forEach(button => {
            button.onclick = function() {
                const columnName = this.textContent.trim();
                // 查找匹配的列索引
                const index = headers.findIndex(header =>
                    header.textContent.trim().includes(columnName) ||
                    columnName.includes(header.textContent.trim())
                );

                if (index !== -1) {
                    // 切换列的显示状态
                    const cells = table.querySelectorAll(`tr > *:nth-child(${index+1})`);
                    const isVisible = cells[0].style.display !== 'none';

                    cells.forEach(cell => {
                        cell.style.display = isVisible ? 'none' : '';
                    });

                    // 更新按钮样式
                    this.style.backgroundColor = isVisible ? '#f0f0f0' : '#e0e0ff';
                }
            };
        });
    }
};

// 修改左侧菜单点击事件
document.querySelectorAll("li[data-url]").forEach(item => {
    item.addEventListener("click", function (e) {
        e.stopPropagation();
        const iframe = document.getElementById("content-frame");
        const placeholder = document.getElementById("placeholder");
        const allPanels = document.querySelectorAll(".analysis-panel");

        // 隐藏所有分析区
        allPanels.forEach(p => p.style.display = "none");

        // 设置iframe内容
        const url = this.getAttribute("data-url");
        iframe.src = url;
        iframe.style.display = "block";
        placeholder.style.display = "none";

        // 显示对应分析区
        const panelId = this.getAttribute("data-panel");
        const panel = document.getElementById(panelId);
        if (panel) {
            panel.style.display = "block";
            // 重置按钮样式
            panel.querySelectorAll('button').forEach(btn => {
                btn.style.backgroundColor = '#fafafa';
            });
        }
    });
});
// static/main.js

document.addEventListener('DOMContentLoaded', () => {
    const navLinks = document.querySelectorAll('#nav-menu a');
    const tableTitle = document.getElementById('table-title');
    const tableContainer = document.getElementById('table-container');
    const columnControls = document.getElementById('column-controls');

    // 为每个导航链接添加点击事件
    navLinks.forEach(link => {
        link.addEventListener('click', (event) => {
            event.preventDefault();
            const tableId = event.target.getAttribute('data-table-id');
            if (tableId) {
                // 移除所有链接的active类
                navLinks.forEach(l => l.classList.remove('active'));
                // 为当前链接添加active类
                event.target.classList.add('active');
                // 加载表格数据
                loadTableData(tableId);
            }
        });
    });

    // 函数：从后端加载数据并渲染
    async function loadTableData(tableId) {
        try {
            const response = await fetch(`/get_table_data/${tableId}`);
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            const tableData = await response.json();

            if (tableData.error) {
                throw new Error(tableData.error);
            }

            renderTable(tableData);
            renderColumnControls(tableData.headers);

        } catch (error) {
            tableTitle.textContent = '加载失败';
            tableContainer.innerHTML = `<p style="color: red;">错误: ${error.message}</p>`;
            columnControls.innerHTML = '';
            console.error('Failed to load table data:', error);
        }
    }

    // 函数：渲染表格
    function renderTable({ title, headers, data }) {
        // 更新标题
        tableTitle.textContent = title;

        // 清空旧表格
        tableContainer.innerHTML = '';

        if (!data || data.length === 0) {
            tableContainer.innerHTML = '<p>该表格没有数据。</p>';
            return;
        }

        // 创建表格元素
        const table = document.createElement('table');
        table.className = 'data-table';

        // 创建表头
        const thead = document.createElement('thead');
        const headerRow = document.createElement('tr');
        headers.forEach(headerText => {
            const th = document.createElement('th');
            th.textContent = headerText;
            headerRow.appendChild(th);
        });
        thead.appendChild(headerRow);
        table.appendChild(thead);

        // 创建表体
        const tbody = document.createElement('tbody');
        data.forEach(rowData => {
            const tr = document.createElement('tr');
            headers.forEach(header => {
                const td = document.createElement('td');
                td.textContent = rowData[header] || ''; // 处理可能的空值
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });
        table.appendChild(tbody);

        // 将表格添加到容器
        tableContainer.appendChild(table);
    }

    // 函数：渲染右侧列控制按钮
    function renderColumnControls(headers) {
        // 清空旧的控制按钮
        columnControls.innerHTML = '';

        headers.forEach(header => {
            const button = document.createElement('button');
            button.textContent = header;
            button.className = 'column-btn active'; // 默认所有列都显示
            button.dataset.columnName = header; // 用data属性存储列名

            // 添加点击事件来切换列的可见性
            button.addEventListener('click', () => toggleColumnVisibility(header, button));

            columnControls.appendChild(button);
        });
    }

    // 函数：切换列的可见性
    function toggleColumnVisibility(columnName, button) {
        const table = document.querySelector('.data-table');
        if (!table) return;

        // 找到该列在表头中的索引
        const headerCells = table.querySelectorAll('th');
        let columnIndex = -1;
        headerCells.forEach((th, index) => {
            if (th.textContent.trim() === columnName) {
                columnIndex = index;
            }
        });

        if (columnIndex === -1) return;

        // 切换该列所有单元格（包括表头）的显示状态
        const rows = table.querySelectorAll('tr');
        rows.forEach(row => {
            const cell = row.cells[columnIndex];
            if (cell) {
                // 检查当前状态并切换
                const isHidden = cell.style.display === 'none';
                cell.style.display = isHidden ? '' : 'none'; // ''恢复默认值
            }
        });

        // 切换按钮的active状态以提供视觉反馈
        button.classList.toggle('active');
    }
});
