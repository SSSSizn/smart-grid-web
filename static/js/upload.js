document.addEventListener('DOMContentLoaded', function() {
    const uploadArea = document.getElementById('uploadArea');
    const fileInput = document.getElementById('fileInput');
    const fileList = document.getElementById('fileList');
    const startUploadBtn = document.getElementById('startUploadBtn');
    const uploadProgress = document.getElementById('uploadProgress');
    const progressBar = document.getElementById('progressBar');
    const progressText = document.getElementById('progressText');
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

    // 处理选择的文件
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
                addFileToUI(file, selectedFiles.length - 1);
            }
        }

        updateUploadButtonState();
    }

    // 添加文件到UI
    function addFileToUI(file, index) {
        const fileItem = document.createElement('div');
        fileItem.className = 'file-item';
        fileItem.innerHTML = `
            <span>${file.name} (${formatFileSize(file.size)})</span>
            <button class="file-remove" data-index="${index}">&times;</button>
        `;
        fileList.appendChild(fileItem);

        // 添加删除事件
        fileItem.querySelector('.file-remove').addEventListener('click', function() {
            const idx = parseInt(this.getAttribute('data-index'));
            selectedFiles.splice(idx, 1);
            fileList.innerHTML = '';
            selectedFiles.forEach((file, i) => addFileToUI(file, i));
            updateUploadButtonState();
        });
    }

    // 更新上传按钮状态
    function updateUploadButtonState() {
        startUploadBtn.disabled = selectedFiles.length === 0;
    }

    // 格式化文件大小
    function formatFileSize(bytes) {
        if (bytes < 1024) return bytes + ' B';
        else if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
        else return (bytes / 1048576).toFixed(1) + ' MB';
    }

    // 显示消息
    function showMessage(text, type) {
        message.textContent = text;
        message.className = 'message ' + type;
        setTimeout(() => {
            message.className = 'message';
        }, 3000);
    }

    // 开始上传
    startUploadBtn.addEventListener('click', function() {
        if (selectedFiles.length === 0) return;

        const formData = new FormData();
        selectedFiles.forEach(file => {
            formData.append('files', file);
        });

        uploadProgress.style.display = 'block';
        progressBar.style.width = '0%';
        progressText.textContent = '0%';

        fetch('/upload', {
            method: 'POST',
            body: formData,
            onUploadProgress: function(e) {
                if (e.lengthComputable) {
                    const percent = Math.round((e.loaded / e.total) * 100);
                    progressBar.style.width = percent + '%';
                    progressText.textContent = percent + '%';
                }
            }
        })
        .then(response => {
            if (!response.ok) {
                throw new Error('上传失败');
            }
            return response.json();
        })
        .then(data => {
            if (data.success) {
                showMessage('文件上传成功', 'success');
                // 清空选择的文件
                selectedFiles = [];
                fileList.innerHTML = '';
                updateUploadButtonState();
                // 3秒后跳转到数据表页面
                setTimeout(() => {
                    window.location.href = '/all_sheets';
                }, 1000);
            } else {
                showMessage(data.message || '上传失败', 'error');
            }
        })
        .catch(error => {
            showMessage(error.message, 'error');
        })
        .finally(() => {
            setTimeout(() => {
                uploadProgress.style.display = 'none';
            }, 1000);
        });
    });
});