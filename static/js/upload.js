// 修改 smart-grid-web/static/js/upload.js
document.addEventListener('DOMContentLoaded', function() {
    const uploadArea = document.getElementById('uploadArea');
    const fileInput = document.getElementById('fileInput');
    const startUploadBtn = document.getElementById('startUploadBtn');
    const message = document.getElementById('message');

    let selectedFiles = [];

    // 点击上传区域触发文件选择
    uploadArea.addEventListener('click', function() {
        fileInput.click();
    });

    // 拖拽功能
    uploadArea.addEventListener('dragover', function(e) {
        e.preventDefault();
        uploadArea.style.borderColor = '#3498db';
        uploadArea.style.backgroundColor = '#f0f8ff';
    });

    uploadArea.addEventListener('dragleave', function() {
        uploadArea.style.borderColor = '#ccc';
        uploadArea.style.backgroundColor = 'white';
    });

    uploadArea.addEventListener('drop', function(e) {
        e.preventDefault();
        uploadArea.style.borderColor = '#ccc';
        uploadArea.style.backgroundColor = 'white';

        if (e.dataTransfer.files.length) {
            handleFiles(e.dataTransfer.files);
        }
    });

    // 文件选择变化时处理
    fileInput.addEventListener('change', function() {
        if (this.files.length) {
            handleFiles(this.files);
        }
    });

    // 处理选择的文件 - 仅做简单验证，不实际上传
    function handleFiles(files) {
        for (let i = 0; i < files.length; i++) {
            const file = files[i];

            // 检查文件类型
            if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
                showMessage('请上传Excel文件 (.xlsx, .xls)', 'error');
                continue;
            }

            // 检查是否已选择该文件
            if (!selectedFiles.some(f => f.name === file.name && f.size === file.size)) {
                selectedFiles.push(file);
            }
        }

        updateUploadButtonState();

        // 选择文件后自动启用上传按钮并显示提示
        if (selectedFiles.length > 0) {
            showMessage(`已选择 ${selectedFiles.length} 个文件，点击上传按钮继续`, 'success');
        }
    }

    // 更新上传按钮状态
    function updateUploadButtonState() {
        startUploadBtn.disabled = selectedFiles.length === 0;
    }

    // 显示消息
    function showMessage(text, type) {
        message.textContent = text;
        message.className = 'message ' + type;
        setTimeout(() => {
            message.className = 'message';
        }, 3000);
    }

    // 开始上传按钮点击事件 - 不实际上传，直接跳转
    startUploadBtn.addEventListener('click', function() {
        if (selectedFiles.length === 0) return;

        showMessage('文件正在处理中，即将跳转到数据表页面...', 'success');

        // 1秒后跳转到数据表页面
        setTimeout(() => {
            window.location.href = '/version1';
        }, 1000);
    });
});