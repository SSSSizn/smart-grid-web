// 点击展开/收起子菜单
document.querySelectorAll(".has-submenu").forEach(item => {
  item.addEventListener("click", function (e) {
    e.stopPropagation();
    this.classList.toggle("active");
  });
});

// 点击“配电线路负载情况”显示右侧内容
document.getElementById("line-load").addEventListener("click", function (e) {
  e.stopPropagation();

  const iframe = document.getElementById("content-frame");
  const placeholder = document.getElementById("placeholder");
  const analysisPanel = document.getElementById("analysis-panel");

  // 设置iframe显示内容
  iframe.src = "/line-load";  // 👈 这是你第一个页面的文件名
  iframe.style.display = "block";
  placeholder.style.display = "none";

  // 显示右侧分析区
  analysisPanel.style.display = "block";
});
