document.addEventListener('DOMContentLoaded', function() {
    loadSheetList();
});

// 加载所有sheet名称
function loadSheetList() {
    fetch('/api/sheets')
        .then(response => response.json())
        .then(sheetNames => {
            const sheetList = document.getElementById('sheetList');
            sheetList.innerHTML = '';

            sheetNames.forEach((name, index) => {
                const li = document.createElement('li');
                li.className = 'sheet-item';
                li.textContent = name;
                li.onclick = () => loadSheetData(name, li);
                sheetList.appendChild(li);

                // 默认加载第一个sheet
                if (index === 0) {
                    loadSheetData(name, li);
                }
            });
        })
        .catch(error => {
            console.error('Error loading sheet list:', error);
            alert('加载sheet列表失败');
        });
}

// 加载指定sheet的数据
function loadSheetData(sheetName, element) {
    // 更新活动状态
    document.querySelectorAll('.sheet-item').forEach(item => {
        item.classList.remove('active');
    });
    element.classList.add('active');

    // 显示加载状态
    document.getElementById('loading').style.display = 'flex';
    document.getElementById('tableContainer').style.display = 'none';
    document.getElementById('sheetInfo').style.display = 'none';

    fetch(`/api/sheet/${encodeURIComponent(sheetName)}`)
        .then(response => response.json())
        .then(data => {
            if (data.error) {
                throw new Error(data.error);
            }

            // 显示sheet信息
            document.getElementById('sheetName').textContent = sheetName;
            document.getElementById('sheetDescription').textContent = `共 ${data.length} 条记录`;
            document.getElementById('sheetInfo').style.display = 'block';

            // 渲染表格
            renderTable(data);

            // 隐藏加载状态，显示表格
            document.getElementById('loading').style.display = 'none';
            document.getElementById('tableContainer').style.display = 'block';
        })
        .catch(error => {
            console.error('Error loading sheet data:', error);
            alert('加载数据失败: ' + error.message);
            document.getElementById('loading').style.display = 'none';
        });
}

// 渲染表格
function renderTable(data) {
    if (!data || data.length === 0) {
        document.getElementById('tableHead').innerHTML = '<tr><th>无数据</th></tr>';
        document.getElementById('tableBody').innerHTML = '<tr><td>该数据表没有内容</td></tr>';
        return;
    }

    // 获取所有列名
    const columns = Object.keys(data[0]);

    // 生成表头
    const headerRow = document.createElement('tr');
    columns.forEach(column => {
        const th = document.createElement('th');
        th.textContent = column;
        headerRow.appendChild(th);
    });
    document.getElementById('tableHead').innerHTML = '';
    document.getElementById('tableHead').appendChild(headerRow);

    // 生成表体
    const tableBody = document.getElementById('tableBody');
    tableBody.innerHTML = '';

    data.forEach(row => {
        const tr = document.createElement('tr');
        columns.forEach(column => {
            const td = document.createElement('td');
            td.textContent = row[column] || '';
            tr.appendChild(td);
        });
        tableBody.appendChild(tr);
    });
}
