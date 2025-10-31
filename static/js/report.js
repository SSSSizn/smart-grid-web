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
        this.initNavigation();
        this.initResizer();
        this.initTextEditing();
    }

    // 初始化导航栏功能
    initNavigation() {
        // 点击展开/收起子菜单
        document.querySelectorAll(".has-submenu").forEach(item => {
            item.addEventListener("click", function (e) {
                if (!e.target.matches('input[type="checkbox"]') && !e.target.matches('label')) {
                    e.stopPropagation();
                    this.classList.toggle("active");
                }
            });
        });

        // 导航搜索功能
        document.getElementById('navSearch').addEventListener('input', (e) => {
            this.filterNavigation(e.target.value);
        });

        // 导航复选框选择事件
        document.querySelectorAll('.table-checkbox').forEach(checkbox => {
            checkbox.addEventListener('change', (e) => {
                this.handleTableSelection(e.target);
            });
        });

        // 全选按钮
        document.getElementById('selectAll').addEventListener('click', () => {
            this.selectAllTables();
        });

        // 清除选择按钮
        document.getElementById('clearSelection').addEventListener('click', () => {
            this.clearAllSelections();
        });
    }

    // 初始化拖拽调整大小
    initResizer() {
        const resizer = document.getElementById('sidebarResizer');
        const sidebar = document.querySelector('.combined-sidebar');
        let isResizing = false;

        resizer.addEventListener('mousedown', (e) => {
            isResizing = true;
            document.body.style.cursor = 'col-resize';
            document.body.style.userSelect = 'none';

            const startX = e.clientX;
            const startWidth = parseInt(document.defaultView.getComputedStyle(sidebar).width, 10);

            const onMouseMove = (e) => {
                if (!isResizing) return;
                const width = startWidth + e.clientX - startX;
                if (width > 250 && width < 600) {
                    sidebar.style.width = width + 'px';
                    sidebar.style.flex = `0 0 ${width}px`;
                }
            };

            const onMouseUp = () => {
                isResizing = false;
                document.body.style.cursor = '';
                document.body.style.userSelect = '';
                document.removeEventListener('mousemove', onMouseMove);
                document.removeEventListener('mouseup', onMouseUp);
            };

            document.addEventListener('mousemove', onMouseMove);
            document.addEventListener('mouseup', onMouseUp);
        });

        // 防止拖拽时选择文本
        resizer.addEventListener('dragstart', (e) => {
            e.preventDefault();
        });
    }

    // 选择所有表格
    selectAllTables() {
        const checkboxes = document.querySelectorAll('.table-checkbox');
        const allChecked = Array.from(checkboxes).every(checkbox => checkbox.checked);

        checkboxes.forEach(checkbox => {
            checkbox.checked = !allChecked;
            this.handleTableSelection(checkbox);
        });

        // 更新按钮文本
        const selectAllBtn = document.getElementById('selectAll');
        selectAllBtn.innerHTML = allChecked ?
            '<i class="fas fa-check-double"></i>' :
            '<i class="fas fa-times"></i>';
        selectAllBtn.title = allChecked ? '全选' : '取消全选';
    }

    // 处理表格选择
    handleTableSelection(checkbox) {
        const tableId = checkbox.value;

        if (checkbox.checked) {
            this.selectedTables.add(tableId);
            checkbox.closest('.nav-table-item').classList.add('selected');
        } else {
            this.selectedTables.delete(tableId);
            checkbox.closest('.nav-table-item').classList.remove('selected');
        }

        this.updateSelectedCount();
        this.updateGenerateButton();
    }

    // 过滤导航栏内容
    filterNavigation(searchTerm) {
        const searchLower = searchTerm.toLowerCase().trim();
        if (!searchLower) {
            // 显示所有项目
            document.querySelectorAll('.nav-tree li').forEach(item => {
                item.style.display = 'block';
            });
            return;
        }

        const navItems = document.querySelectorAll('.nav-tree li');
        let foundItems = new Set();

        // 第一遍：标记匹配的项目
        navItems.forEach(item => {
            const text = item.textContent.toLowerCase();
            if (text.includes(searchLower) && item.classList.contains('nav-table-item')) {
                item.style.display = 'block';
                foundItems.add(item);

                // 显示所有父级菜单
                let parent = item.parentElement;
                while (parent && parent.classList.contains('submenu')) {
                    parent.style.display = 'block';
                    const parentLi = parent.parentElement;
                    parentLi.classList.add('active');
                    foundItems.add(parentLi);
                    parent = parentLi.parentElement;
                }
            } else {
                item.style.display = 'none';
            }
        });

        // 第二遍：确保所有相关父级都显示
        foundItems.forEach(item => {
            item.style.display = 'block';
        });
    }

    // 清除所有选择
    clearAllSelections() {
        document.querySelectorAll('.table-checkbox').forEach(checkbox => {
            checkbox.checked = false;
            this.handleTableSelection(checkbox);
        });
    }

    async loadTableConfigs() {
        try {
            const response = await fetch('/api/table-configs');
            const data = await response.json();

            if (data.success) {
                this.tableConfigs = data.data;
                this.syncNavigationWithConfigs();
            } else {
                this.showError('加载表格配置失败: ' + data.message);
            }
        } catch (error) {
            this.showError('加载表格配置时出错: ' + error.message);
        }
    }

    // 同步导航栏与表格配置
    syncNavigationWithConfigs() {
        this.tableConfigs.forEach(table => {
            const checkbox = document.querySelector(`.table-checkbox[value="${table.id}"]`);
            if (checkbox) {
                // 确保标签文本与配置一致
                const label = checkbox.nextElementSibling;
                if (label && label.textContent !== table.title) {
                    label.textContent = table.title;
                }
            }
        });
    }

    updateSelectedCount() {
        const countElement = document.getElementById('selectedCount');
        countElement.textContent = this.selectedTables.size;
    }

    // 初始化文字编辑功能
    initTextEditing() {
        // 使用事件委托来处理动态生成的编辑按钮
        document.getElementById('reportPreview').addEventListener('click', (e) => {
            // 编辑按钮点击
            if (e.target.closest('.edit-description-btn')) {
                const btn = e.target.closest('.edit-description-btn');
                const tableId = btn.closest('.editable-text').dataset.tableId;
                this.enableTextEditing(tableId);
            }

            // 保存按钮点击
            if (e.target.closest('.save-description')) {
                const btn = e.target.closest('.save-description');
                const tableId = btn.dataset.tableId;
                this.saveTextContent(tableId);
            }

            // 取消按钮点击
            if (e.target.closest('.cancel-edit')) {
                const btn = e.target.closest('.cancel-edit');
                const tableId = btn.closest('.editable-text').dataset.tableId;
                this.cancelTextEditing(tableId);
            }
        });
    }

    // 启用文字编辑
    enableTextEditing(tableId) {
        const container = document.querySelector(`.editable-text[data-table-id="${tableId}"]`);
        if (!container) return;

        const contentDiv = container.querySelector('.description-content');
        const editControls = container.querySelector('.edit-controls');
        const editBtn = container.querySelector('.edit-description-btn');
        const textarea = container.querySelector('.description-edit');

        // 保存原始内容以便取消时恢复
        container.dataset.originalContent = contentDiv.innerHTML;

        // 切换到编辑模式
        contentDiv.style.display = 'none';
        editControls.style.display = 'block';
        editBtn.style.display = 'none';

        // 设置文本区域内容
        textarea.value = contentDiv.textContent || '';
        textarea.focus();
    }

    // 取消文字编辑
    cancelTextEditing(tableId) {
        const container = document.querySelector(`.editable-text[data-table-id="${tableId}"]`);
        if (!container) return;

        const contentDiv = container.querySelector('.description-content');
        const editControls = container.querySelector('.edit-controls');
        const editBtn = container.querySelector('.edit-description-btn');

        // 恢复原始内容
        contentDiv.innerHTML = container.dataset.originalContent || '';

        // 切换回显示模式
        contentDiv.style.display = 'block';
        editControls.style.display = 'none';
        editBtn.style.display = 'block';
    }

    // 保存文字内容
    async saveTextContent(tableId) {
        const container = document.querySelector(`.editable-text[data-table-id="${tableId}"]`);
        if (!container) return;

        const contentDiv = container.querySelector('.description-content');
        const editControls = container.querySelector('.edit-controls');
        const editBtn = container.querySelector('.edit-description-btn');
        const textarea = container.querySelector('.description-edit');

        const newContent = textarea.value.trim();

        try {
            const response = await fetch('/api/save-table-description', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    table_id: tableId,
                    content: newContent
                })
            });

            const data = await response.json();

            if (data.success) {
                this.showSuccess('描述内容保存成功！');

                // 更新显示的内容
                contentDiv.innerHTML = newContent || '暂无内容描述。';

                // 切换回显示模式
                contentDiv.style.display = 'block';
                editControls.style.display = 'none';
                editBtn.style.display = 'block';
            } else {
                this.showError('保存失败: ' + data.message);
            }
        } catch (error) {
            this.showError('保存时出错: ' + error.message);
        }
    }

    // 在 displayReportPreview 方法中调用初始化文字编辑
    displayReportPreview(htmlContent) {
        const preview = document.getElementById('reportPreview');
        this.currentHtmlContent = htmlContent;

        // 移除空预览状态
        preview.innerHTML = htmlContent;

        // 确保预览区域滚动到顶部
        preview.scrollTop = 0;

        // 初始化文字编辑功能
        this.initTextEditing();

        // 显示统计信息
        this.showReportStats();
    }
    updateGenerateButton() {
        const generateBtn = document.getElementById('generateReport');
        const downloadBtn = document.getElementById('downloadReport');
        const exportPreviewBtn = document.getElementById('exportPreview');

        const hasSelection = this.selectedTables.size > 0;

        generateBtn.disabled = !hasSelection;
        downloadBtn.disabled = !hasSelection;
        exportPreviewBtn.disabled = !hasSelection;
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
                this.updateReportStats(data.table_count);
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

        // 重新初始化文字编辑功能（因为内容已更新）
        this.initTextEditing();

        // 显示统计信息
        this.showReportStats();
    }

    updateReportStats(tableCount) {
        const stats = document.getElementById('reportStats');
        stats.style.display = 'block';

        const count = tableCount || this.selectedTables.size;
        document.getElementById('tableCount').textContent = count;
        document.getElementById('dataCount').textContent = (count * 42).toLocaleString();
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
                ${document.getElementById('reportPreview').innerHTML}
            </body>
            </html>
        `);
        printWindow.document.close();
        printWindow.focus();

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