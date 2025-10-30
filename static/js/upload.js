// 智能电网分析系统 - 文件上传功能（修复版本）
document.addEventListener('DOMContentLoaded', function() {
    // DOM 元素
    const uploadArea = document.getElementById('uploadArea');
    const fileInput = document.getElementById('fileInput');
    const startUploadBtn = document.getElementById('startUploadBtn');
    const message = document.getElementById('message');
    const fileInfo = document.getElementById('fileInfo');
    const fileName = document.getElementById('fileName');
    const fileSize = document.getElementById('fileSize');
    const fileRemoveBtn = document.getElementById('fileRemoveBtn');

    // 状态变量
    let selectedFile = null;
    let isUploading = false;
    let cancelBtn = null;
    let isProgrammaticNavigation = false; // 新增：标记是否为程序化跳转

    // 点击上传区域触发文件选择
    uploadArea.addEventListener('click', function() {
        if (!isUploading) {
            fileInput.click();
        }
    });

    // 拖拽功能
    uploadArea.addEventListener('dragover', function(e) {
        e.preventDefault();
        if (!isUploading) {
            uploadArea.classList.add('dragover');
        }
    });

    uploadArea.addEventListener('dragleave', function() {
        uploadArea.classList.remove('dragover');
    });

    uploadArea.addEventListener('drop', function(e) {
        e.preventDefault();
        uploadArea.classList.remove('dragover');

        if (!isUploading && e.dataTransfer.files.length) {
            handleFile(e.dataTransfer.files[0]);
        }
    });

    // 文件选择变化时处理
    fileInput.addEventListener('change', function() {
        if (this.files.length && !isUploading) {
            handleFile(this.files[0]);
        }
    });

    // 处理选择的文件
    function handleFile(file) {
        // 检查文件类型
        if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
            showMessage('请上传Excel文件 (.xlsx, .xls)', 'error');
            return;
        }

        // 检查文件大小 (限制为50MB)
        const maxSize = 50 * 1024 * 1024; // 50MB
        if (file.size > maxSize) {
            showMessage('文件大小不能超过50MB', 'error');
            return;
        }

        // 设置选中的文件
        selectedFile = file;
        updateFileInfo();
        updateUploadButtonState();
        showMessage('文件已选择，点击上传按钮开始上传', 'success');
    }

    // 更新文件信息显示
    function updateFileInfo() {
        if (selectedFile) {
            fileName.textContent = selectedFile.name;
            fileSize.textContent = formatFileSize(selectedFile.size);
            fileInfo.style.display = 'block';
            uploadArea.classList.add('active');
        } else {
            fileInfo.style.display = 'none';
            uploadArea.classList.remove('active');
        }
    }

    // 文件大小格式化
    function formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    // 移除文件
    fileRemoveBtn.addEventListener('click', function(e) {
        e.stopPropagation();
        resetUploadState();
        showMessage('文件已移除', 'info');
    });

    // 重置上传状态
    function resetUploadState() {
        isUploading = false;
        selectedFile = null;
        fileInput.value = '';
        updateFileInfo();
        updateUploadButtonState();
        if (cancelBtn) cancelBtn.style.display = 'none';
    }

    // 更新上传按钮状态
    function updateUploadButtonState() {
        startUploadBtn.disabled = !selectedFile || isUploading;

        if (isUploading) {
            startUploadBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i><span>上传中...</span>';
        } else {
            startUploadBtn.innerHTML = '<i class="fas fa-upload"></i><span>开始上传</span>';
        }
    }

    // 显示消息
    function showMessage(text, type) {
        message.textContent = text;
        message.className = 'message ' + type;
        message.style.display = 'block';

        // 自动隐藏非错误消息
        if (type !== 'error') {
            setTimeout(() => {
                message.style.display = 'none';
            }, 5000);
        }
    }

    // 文件上传函数 - 修复版本
    async function uploadFile(file) {
        try {
            const formData = new FormData();
            formData.append('file', file);

            console.log('开始上传文件:', file.name, '大小:', file.size);

            // 添加超时控制
            const controller = new AbortController();
            const timeoutId = setTimeout(() => controller.abort(), 30000); // 30秒超时

            const response = await fetch('/upload', {
                method: 'POST',
                body: formData,
                signal: controller.signal,
            });

            clearTimeout(timeoutId);

            console.log('服务器响应状态:', response.status);

            if (!response.ok) {
                let errorMessage = `HTTP错误: ${response.status}`;
                try {
                    const errorData = await response.json();
                    errorMessage = errorData.error || errorMessage;
                } catch (e) {
                    // 如果响应不是JSON，尝试读取文本
                    try {
                        const text = await response.text();
                        if (text) errorMessage = text;
                    } catch (textError) {
                        console.error('读取错误响应失败:', textError);
                    }
                }
                throw new Error(errorMessage);
            }

            const result = await response.json();
            console.log('上传成功:', result);
            return result;

        } catch (error) {
            console.error('上传过程错误:', error);
            if (error.name === 'AbortError') {
                throw new Error('上传超时，请检查网络连接或稍后重试');
            }
            throw error;
        }
    }

    // 开始上传按钮点击事件 - 修复版本
    startUploadBtn.addEventListener('click', async function() {
        if (!selectedFile || isUploading) return;

        isUploading = true;
        updateUploadButtonState();
        if (cancelBtn) cancelBtn.style.display = 'flex';
        showMessage('正在上传文件，请稍候...', 'info');

        try {
            const result = await uploadFile(selectedFile);

            if (result.success) {
                showMessage(`文件上传成功: ${result.filename}`, 'success');
                if (cancelBtn) cancelBtn.style.display = 'none';

                // 标记为程序化跳转，避免触发beforeunload
                isProgrammaticNavigation = true;

                // 上传成功后跳转到表格分析页面
                setTimeout(() => {
                    isUploading = false; // 上传已完成，重置状态
                    updateUploadButtonState();
                    window.location.href = '/version1';
                }, 1500);
            } else {
                showMessage(`上传失败: ${result.error || '未知错误'}`, 'error');
                resetUploadState();
            }
        } catch (error) {
            console.error('上传错误:', error);
            showMessage(`上传失败: ${error.message || '网络错误'}`, 'error');
            resetUploadState();
        }
    });

    // 添加强制取消上传按钮
    function addCancelButton() {
        const cancelBtn = document.createElement('button');
        cancelBtn.innerHTML = '<i class="fas fa-times"></i><span>取消上传</span>';
        cancelBtn.className = 'cancel-upload-btn';
        cancelBtn.style.display = 'none';
        cancelBtn.onclick = function() {
            resetUploadState();
            showMessage('上传已取消', 'info');
        };

        startUploadBtn.parentNode.insertBefore(cancelBtn, startUploadBtn.nextSibling);
        return cancelBtn;
    }

    // 改进的beforeunload处理
    function handleBeforeUnload(event) {
        // 只有在实际上传中且不是程序化跳转时才提示用户
        if (isUploading && !isProgrammaticNavigation) {
            event.preventDefault();
            event.returnValue = '文件正在上传中，确定要离开吗？';
            return '文件正在上传中，确定要离开吗？';
        }
    }

    // 初始化
    function init() {
        updateFileInfo();
        updateUploadButtonState();
        cancelBtn = addCancelButton();

        // 添加页面卸载保护 - 使用改进的处理函数
        window.addEventListener('beforeunload', handleBeforeUnload);

        console.log('文件上传系统初始化完成');
    }

    init();
});