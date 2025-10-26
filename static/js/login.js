document.getElementById('loginForm').addEventListener('submit', function(e) {
    e.preventDefault();

    const username = document.getElementById('username').value;
    const password = document.getElementById('password').value;
    const errorMessage = document.getElementById('errorMessage');

    // 简单的前端验证
    if (!username || !password) {
        errorMessage.textContent = '请输入用户名和密码';
        return;
    }

    // 发送登录请求
    fetch('/login', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify({ username, password })
    })
    .then(response => {
        if (!response.ok) {
            return response.json().then(data => {
                throw new Error(data.message || '登录失败');
            });
        }
        return response.json();
    })
    .then(data => {
        if (data.success) {
            // 登录成功，跳转到首页
            window.location.href = '/upload';
        } else {
            errorMessage.textContent = data.message || '登录失败';
        }
    })
    .catch(error => {
        errorMessage.textContent = error.message;
    });
});