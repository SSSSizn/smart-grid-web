class ReportManager {
    constructor() {
        this.selectedTables = new Set();
        this.tableConfigs = [];
        this.currentHtmlContent = '';
        this.init();
    }

    init() {
        this.loadTableConfigs();
        this.bindEvents();
    }

    async loadTableConfigs() {
        try {
            const response = await fetch('/api/table-configs');
            const data = await response.json();

            if (data.success) {
                this.tableConfigs = data.data;
                this.renderTableList();
            } else {
                this.showError('加载表格配置失败: ' + data.message);
            }
        } catch (error) {
            this.showError('加载表格配置时出错: ' + error.message);
        }
    }

    renderTableList() {
        const tableList = document.getElementById('tableList');
        tableList.innerHTML = '';

        this.tableConfigs.forEach(table => {
            const tableItem = document.createElement('div');
            tableItem.className = 'table-item';
            tableItem.innerHTML = `
                <div class="form-check">
                    <input class="form-check-input table-checkbox" type="checkbox" 
                           value="${table.id}" id="table-${table.id}">
                    <label class="form-check-label" for="table-${table.id}">
                        ${table.title}
                    </label>
                </div>
            `;
            tableList.appendChild(tableItem);
        });

        // 绑定复选框事件
        document.querySelectorAll('.table-checkbox').forEach(checkbox => {
            checkbox.addEventListener('change', (e) => {
                this.handleTableSelection(e.target);
            });
        });

        // 绑定搜索功能
        document.getElementById('tableSearch').addEventListener('input', (e) => {
            this.filterTables(e.target.value);
        });
    }

    handleTableSelection(checkbox) {
        const tableId = checkbox.value;

        if (checkbox.checked) {
            this.selectedTables.add(tableId);
            checkbox.parentElement.parentElement.classList.add('selected');
        } else {
            this.selectedTables.delete(tableId);
            checkbox.parentElement.parentElement.classList.remove('selected');
        }

        this.updateSelectedCount();
        this.updateGenerateButton();
    }

    filterTables(searchTerm) {
        const tableItems = document.querySelectorAll('.table-item');
        const searchLower = searchTerm.toLowerCase();

        tableItems.forEach(item => {
            const label = item.querySelector('.form-check-label').textContent.toLowerCase();
            if (label.includes(searchLower)) {
                item.style.display = 'block';
            } else {
                item.style.display = 'none';
            }
        });
    }

    updateSelectedCount() {
        const countElement = document.getElementById('selectedCount');
        countElement.textContent = this.selectedTables.size;
    }

    updateGenerateButton() {
        const generateBtn = document.getElementById('generateReport');
        generateBtn.disabled = this.selectedTables.size === 0;
    }

    async generateReport() {
        if (this.selectedTables.size === 0) {
            this.showError('请至少选择一个表格');
            return;
        }

        this.showLoading('生成报告预览中...');

        try {
            const response = await fetch('/api/generate-report-preview', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    tables: Array.from(this.selectedTables)
                })
            });

            const data = await response.json();

            if (data.success) {
                this.showSuccess('报告预览生成成功！');
                this.displayReportPreview(data.html_content);
                this.enableDownload();
                this.updateReportStats(data.table_count);
                this.enableExportPreview();
            } else {
                this.showError('生成报告预览失败: ' + data.message);
            }
        } catch (error) {
            this.showError('生成报告预览时出错: ' + error.message);
        } finally {
            this.hideLoading();
        }
    }

    displayReportPreview(htmlContent) {
        const preview = document.getElementById('reportPreview');
        this.currentHtmlContent = htmlContent;

        // 移除空预览状态
        preview.innerHTML = htmlContent;

        // 确保预览区域滚动到顶部
        preview.scrollTop = 0;

        // 显示统计信息
        this.showReportStats();
    }

    enableDownload() {
        const downloadBtn = document.getElementById('downloadReport');
        downloadBtn.disabled = false;
    }

    enableExportPreview() {
        const exportBtn = document.getElementById('exportPreview');
        exportBtn.disabled = false;
    }

    updateReportStats(tableCount) {
        const stats = document.getElementById('reportStats');
        stats.style.display = 'block';

        const count = tableCount || this.selectedTables.size;
        document.getElementById('tableCount').textContent = count;
        document.getElementById('dataCount').textContent = (count * 42).toLocaleString(); // 模拟数据量
        document.getElementById('pageCount').textContent = Math.ceil(count * 1.2);
        document.getElementById('fileSize').textContent = (count * 85) + ' KB';
    }

    showReportStats() {
        const stats = document.getElementById('reportStats');
        stats.style.display = 'block';
    }

    async downloadReport() {
        const format = document.getElementById('exportFormat').value;

        if (format === 'word') {
            this.showLoading('生成Word报告中...');

            try {
                const response = await fetch('/api/generate-full-report', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                    },
                    body: JSON.stringify({
                        tables: Array.from(this.selectedTables)
                    })
                });

                const data = await response.json();

                if (data.success) {
                    window.location.href = '/download-report';
                } else {
                    this.showError('生成下载文件失败: ' + data.message);
                }
            } catch (error) {
                this.showError('生成下载文件时出错: ' + error.message);
            } finally {
                this.hideLoading();
            }
        } else {
            this.showError('PDF导出功能正在开发中，请先使用Word格式');
        }
    }

    printReport() {
        const preview = document.getElementById('reportPreview');
        if (!this.currentHtmlContent) {
            this.showError('请先生成报告预览');
            return;
        }

        const printWindow = window.open('', '_blank');
        printWindow.document.write(`
            <!DOCTYPE html>
            <html>
            <head>
                <title>智能电网分析报告</title>
                <meta charset="UTF-8">
                <style>
                    body { 
                        font-family: "Microsoft YaHei", Arial, sans-serif; 
                        margin: 20px; 
                        line-height: 1.6;
                        color: #333;
                    }
                    .report-header { 
                        text-align: center; 
                        margin-bottom: 30px; 
                        padding-bottom: 20px;
                        border-bottom: 2px solid #333;
                    }
                    .report-title { 
                        color: #2c3e50; 
                        font-size: 28px;
                        margin-bottom: 10px;
                    }
                    .report-meta {
                        color: #666;
                        font-size: 14px;
                    }
                    .table-section { 
                        margin-bottom: 30px; 
                        page-break-inside: avoid;
                        break-inside: avoid;
                    }
                    .table-section h3 {
                        color: #2c3e50;
                        border-left: 4px solid #3498db;
                        padding-left: 15px;
                        margin-bottom: 15px;
                    }
                    table { 
                        width: 100%; 
                        border-collapse: collapse; 
                        margin: 15px 0;
                        font-size: 12px;
                    }
                    th, td { 
                        border: 1px solid #ddd; 
                        padding: 8px; 
                        text-align: left; 
                    }
                    th { 
                        background-color: #f2f2f2; 
                        font-weight: bold;
                    }
                    .analysis-summary { 
                        background: #f8f9fa; 
                        padding: 15px; 
                        margin: 10px 0;
                        border-left: 4px solid #3498db;
                    }
                    .analysis-summary h4 {
                        margin-top: 0;
                    }
                    .report-summary {
                        margin-top: 30px;
                        padding: 20px;
                        background: #f8f9fa;
                        border-top: 2px solid #ddd;
                    }
                    .signature {
                        text-align: right;
                        margin-top: 30px;
                        font-style: italic;
                        color: #666;
                    }
                    @media print {
                        .page-break { 
                            page-break-after: always;
                        }
                        body {
                            margin: 0.5in;
                        }
                    }
                </style>
            </head>
            <body>
                ${preview.innerHTML}
            </body>
            </html>
        `);
        printWindow.document.close();
        printWindow.focus();

        // 等待内容加载完成后打印
        setTimeout(() => {
            printWindow.print();
            printWindow.close();
        }, 500);
    }

    exportPreview() {
        if (!this.currentHtmlContent) {
            this.showError('请先生成报告预览');
            return;
        }

        // 创建可下载的HTML文件
        const blob = new Blob([this.currentHtmlContent], { type: 'text/html' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = '智能电网分析报告预览.html';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        this.showSuccess('预览文件导出成功！');
    }

    bindEvents() {
        // 全选/全不选
        document.getElementById('selectAll').addEventListener('change', (e) => {
            const checkboxes = document.querySelectorAll('.table-checkbox');
            checkboxes.forEach(checkbox => {
                checkbox.checked = e.target.checked;
                this.handleTableSelection(checkbox);
            });
        });

        // 生成报告按钮
        document.getElementById('generateReport').addEventListener('click', () => {
            this.generateReport();
        });

        // 下载报告按钮
        document.getElementById('downloadReport').addEventListener('click', () => {
            this.downloadReport();
        });

        // 打印报告按钮
        document.getElementById('printReport').addEventListener('click', () => {
            this.printReport();
        });

        // 导出预览按钮
        document.getElementById('exportPreview').addEventListener('click', () => {
            this.exportPreview();
        });
    }

    showLoading(message) {
        const loadingModal = new bootstrap.Modal(document.getElementById('loadingModal'));
        document.getElementById('loadingMessage').textContent = message;
        loadingModal.show();
    }

    hideLoading() {
        const loadingModal = bootstrap.Modal.getInstance(document.getElementById('loadingModal'));
        if (loadingModal) {
            loadingModal.hide();
        }
    }

    showSuccess(message) {
        this.showAlert(message, 'success');
    }

    showError(message) {
        this.showAlert(message, 'danger');
    }

    showAlert(message, type) {
        // 移除现有的警告
        const existingAlerts = document.querySelectorAll('.alert');
        existingAlerts.forEach(alert => alert.remove());

        const alert = document.createElement('div');
        alert.className = `alert alert-${type} alert-dismissible fade show position-fixed`;
        alert.style.cssText = 'top: 80px; right: 20px; z-index: 1050; min-width: 300px;';
        alert.innerHTML = `
            ${message}
            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
        `;

        document.body.appendChild(alert);

        setTimeout(() => {
            if (alert.parentNode) {
                alert.remove();
            }
        }, 5000);
    }
}

// 初始化报告管理器
document.addEventListener('DOMContentLoaded', function() {
    new ReportManager();
});