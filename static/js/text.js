// ====== 修复：添加变量定义 ======
const editableDiv = document.getElementById('editableText');
const editBtn = document.getElementById('editBtn');
const saveBtn = document.getElementById('saveBtn');
const cancelBtn = document.getElementById('cancelBtn');

let originalText = ''; // 用于保存原始内容
// ===============================

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
    const page_id = getPageId();
    
    try {
        const res = await fetch(`/save/${page_id}`, {
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

// 添加键盘事件支持
editableDiv.addEventListener('keydown', function(event) {
    if (event.key === 'Enter' && !event.shiftKey) {
        event.preventDefault();
        // 可选：按Enter键自动保存
        // saveText();
    }
});


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