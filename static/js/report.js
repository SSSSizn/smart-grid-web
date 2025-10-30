class ReportManager {
    constructor() {
        this.selectedTables = new Set();
        this.tableConfigs = [];
        this.isEditing = false;
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

        // 添加编辑功能
        this.makeContentEditable();

        // 确保预览区域滚动到顶部
        preview.scrollTop = 0;

        // 显示统计信息
        this.showReportStats();
    }

    makeContentEditable() {
        if (!this.isEditing) return;

        const preview = document.getElementById('reportPreview');
        const editableElements = preview.querySelectorAll('h1, h2, h3, h4, h5, p, li, td');

        editableElements.forEach(element => {
            element.classList.add('editable');
            element.setAttribute('contenteditable', 'true');
        });
    }

    enableEditMode() {
        this.isEditing = true;
        const preview = document.getElementById('reportPreview');
        preview.classList.add('editing');

        // 添加编辑工具栏
        this.addEditToolbar();

        // 使内容可编辑
        this.makeContentEditable();

        this.showSuccess('现在可以编辑报告内容了');
    }

    disableEditMode() {
        this.isEditing = false;
        const preview = document.getElementById('reportPreview');
        preview.classList.remove('editing');

        // 移除编辑工具栏
        this.removeEditToolbar();

        // 移除可编辑属性
        const editableElements = preview.querySelectorAll('.editable');
        editableElements.forEach(element => {
            element.classList.remove('editable');
            element.removeAttribute('contenteditable');
        });

        this.showSuccess('已退出编辑模式');
    }

    addEditToolbar() {
        const preview = document.getElementById('reportPreview');

        const toolbar = document.createElement('div');
        toolbar.className = 'edit-toolbar';
        toolbar.innerHTML = `
            <div class="toolbar-buttons">
                <button class="toolbar-btn" data-action="bold">
                    <i class="fas fa-bold"></i> 粗体
                </button>
                <button class="toolbar-btn" data-action="italic">
                    <i class="fas fa-italic"></i> 斜体
                </button>
                <button class="toolbar-btn" data-action="underline">
                    <i class="fas fa-underline"></i> 下划线
                </button>
                <button class="toolbar-btn" data-action="save">
                    <i class="fas fa-save"></i> 保存修改
                </button>
                <button class="toolbar-btn" data-action="cancel">
                    <i class="fas fa-times"></i> 取消编辑
                </button>
            </div>
        `;

        preview.insertBefore(toolbar, preview.firstChild);

        // 绑定工具栏事件
        toolbar.addEventListener('click', (e) => {
            const action = e.target.closest('.toolbar-btn')?.dataset.action;
            if (action) {
                this.handleToolbarAction(action);
            }
        });
    }

    removeEditToolbar() {
        const toolbar = document.querySelector('.edit-toolbar');
        if (toolbar) {
            toolbar.remove();
        }
    }

    handleToolbarAction(action) {
        switch (action) {
            case 'bold':
                document.execCommand('bold');
                break;
            case 'italic':
                document.execCommand('italic');
                break;
            case 'underline':
                document.execCommand('underline');
                break;
            case 'save':
                this.saveEditedContent();
                break;
            case 'cancel':
                this.disableEditMode();
                break;
        }
    }

    async saveEditedContent() {
        const preview = document.getElementById('reportPreview');
        const htmlContent = preview.innerHTML;

        this.showLoading('保存报告中...');

        try {
            const response = await fetch('/api/save-report-content', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    html_content: htmlContent
                })
            });

            const data = await response.json();

            if (data.success) {
                this.showSuccess('报告内容保存成功！');
                this.currentHtmlContent = htmlContent;
            } else {
                this.showError('保存报告内容失败: ' + data.message);
            }
        } catch (error) {
            this.showError('保存报告内容时出错: ' + error.message);
        } finally {
            this.hideLoading();
        }
    }

    enableDownload() {
        const downloadBtn = document.getElementById('downloadReport');
        downloadBtn.disabled = false;
    }

    updateReportStats(tableCount) {
        const stats = document.getElementById('reportStats');
        stats.style.display = 'block';

        const count = tableCount || this.selectedTables.size;
        document.getElementById('tableCount').textContent = count;
        document.getElementById('dataCount').textContent = count * 50; // 模拟数据量
        document.getElementById('pageCount').textContent = Math.ceil(count * 0.5);
        document.getElementById('fileSize').textContent = (count * 100) + ' KB';
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
                        tables: Array.from(this.selectedTables),
                        use_edited_content: this.isEditing
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
                <style>
                    body { font-family: Arial, sans-serif; margin: 20px; }
                    .report-header { text-align: center; margin-bottom: 30px; }
                    .report-title { color: #2c3e50; }
                    .table-section { margin-bottom: 30px; page-break-inside: avoid; }
                    table { width: 100%; border-collapse: collapse; margin: 10px 0; }
                    th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
                    th { background-color: #f2f2f2; }
                    .analysis-summary { background: #f8f9fa; padding: 15px; margin: 10px 0; }
                    @media print {
                        .page-break { page-break-after: always; }
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
        printWindow.print();
        printWindow.close();
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

        // 编辑报告按钮
        document.getElementById('editReport').addEventListener('click', () => {
            if (this.currentHtmlContent) {
                this.enableEditMode();
            } else {
                this.showError('请先生成报告预览');
            }
        });

        // 保存报告按钮
        document.getElementById('saveReport').addEventListener('click', () => {
            if (this.isEditing) {
                this.saveEditedContent();
            } else {
                this.showError('请先进入编辑模式');
            }
        });

        // 打印报告按钮
        document.getElementById('printReport').addEventListener('click', () => {
            this.printReport();
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