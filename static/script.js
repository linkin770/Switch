// 全局变量
let selectedFiles = [];

// 初始化
function init() {
    setupFileInput();
    setupDragAndDrop();
    setupFormatTooltip();
    setupEventListeners();
}

// 设置格式提示
function setupFormatTooltip() {
    const select = document.getElementById('fmt');
    const tooltip = document.getElementById('formatTooltip');
    
    // 更新提示信息
    function updateTooltip() {
        const selectedOption = select.options[select.selectedIndex];
        const supported = selectedOption.dataset.supported;
        tooltip.textContent = `支持的输入格式: ${supported}`;
    }
    
    // 显示提示
    function showTooltip() {
        tooltip.style.opacity = '1';
        tooltip.style.visibility = 'visible';
    }
    
    // 隐藏提示
    function hideTooltip() {
        tooltip.style.opacity = '0';
        tooltip.style.visibility = 'hidden';
    }
    
    // 初始更新
    updateTooltip();
    
    // 当选择变化时更新
    select.addEventListener('change', function() {
        updateTooltip();
        showTooltip();
        // 3秒后自动隐藏
        setTimeout(hideTooltip, 3000);
    });
    
    // 当鼠标悬停在选择器上时显示提示
    select.addEventListener('mouseover', function() {
        updateTooltip();
        showTooltip();
    });
    
    // 当鼠标离开选择器时隐藏提示
    select.addEventListener('mouseout', hideTooltip);
    
    // 当选择器获得焦点时显示提示
    select.addEventListener('focus', function() {
        updateTooltip();
        showTooltip();
    });
    
    // 当选择器失去焦点时隐藏提示
    select.addEventListener('blur', hideTooltip);
}

// 设置文件输入
function setupFileInput() {
    const fileInput = document.getElementById('files');
    fileInput.addEventListener('change', function(e) {
        selectedFiles = Array.from(e.target.files);
        updateFileList();
    });
}

// 设置拖拽功能
function setupDragAndDrop() {
    const fileLabel = document.querySelector('.file-label');
    
    // 拖拽进入
    fileLabel.addEventListener('dragover', function(e) {
        e.preventDefault();
        this.style.borderColor = '#2E8B57';
        this.style.background = 'rgba(46, 139, 87, 0.1)';
    });
    
    // 拖拽离开
    fileLabel.addEventListener('dragleave', function(e) {
        e.preventDefault();
        this.style.borderColor = 'rgba(46, 139, 87, 0.5)';
        this.style.background = 'rgba(46, 139, 87, 0.05)';
    });
    
    // 放置文件
    fileLabel.addEventListener('drop', function(e) {
        e.preventDefault();
        this.style.borderColor = 'rgba(46, 139, 87, 0.5)';
        this.style.background = 'rgba(46, 139, 87, 0.05)';
        
        if (e.dataTransfer.files.length > 0) {
            selectedFiles = Array.from(e.dataTransfer.files);
            document.getElementById('files').files = e.dataTransfer.files;
            updateFileList();
        }
    });
}

// 更新文件列表显示
function updateFileList() {
    const fileInput = document.querySelector('.file-input');
    
    // 移除旧的文件列表
    const oldFileList = document.querySelector('.file-list');
    if (oldFileList) {
        oldFileList.remove();
    }
    
    // 创建新的文件列表
    if (selectedFiles.length > 0) {
        const fileList = document.createElement('div');
        fileList.className = 'file-list';
        fileList.style.marginTop = '10px';
        fileList.style.padding = '10px';
        fileList.style.background = 'rgba(255, 255, 255, 0.5)';
        fileList.style.borderRadius = '8px';
        fileList.style.border = '1px solid rgba(255, 255, 255, 0.3)';
        
        const title = document.createElement('div');
        title.style.fontWeight = '500';
        title.style.marginBottom = '5px';
        title.style.color = '#34495e';
        title.textContent = `已选择 ${selectedFiles.length} 个文件：`;
        fileList.appendChild(title);
        
        selectedFiles.forEach(file => {
            const fileItem = document.createElement('div');
            fileItem.style.padding = '5px 0';
            fileItem.style.borderBottom = '1px solid rgba(255, 255, 255, 0.3)';
            fileItem.style.display = 'flex';
            fileItem.style.justifyContent = 'space-between';
            
            const fileName = document.createElement('span');
            fileName.textContent = file.name;
            
            const fileSize = document.createElement('span');
            fileSize.style.fontSize = '12px';
            fileSize.style.color = '#7f8c8d';
            fileSize.textContent = formatFileSize(file.size);
            
            fileItem.appendChild(fileName);
            fileItem.appendChild(fileSize);
            fileList.appendChild(fileItem);
        });
        
        fileInput.appendChild(fileList);
    }
}

// 格式化文件大小
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// 设置事件监听器
function setupEventListeners() {
    // 开始转换按钮
    document.getElementById('startBtn').addEventListener('click', async function() {
        const resultDiv = document.getElementById('result');
        const downloadsDiv = document.getElementById('downloads');
        resultDiv.innerText = "⏳ 转换中...";
        downloadsDiv.innerHTML = "";

        const form = new FormData();
        form.append('format', document.getElementById('fmt').value);

        const files = document.getElementById('files').files;
        if(files.length===0){
            resultDiv.innerText = "⚠️ 请先选择文件";
            return;
        }
        for (let f of files) form.append('files', f);

        try {
            const res = await fetch('/convert', {method:'POST', body:form});
            const data = await res.json();

            resultDiv.innerText = data.logs || data.msg;

            downloadsDiv.innerHTML = "";
            if(data.success > 0){
                data.logs.split("\n").forEach(line => {
                    if(line.startsWith("✅ 成功：")){
                        const [_, info] = line.split("：");
                        const [oldName, newName] = info.split(" → ");
                        const link = document.createElement("a");
                        link.href = `/outputs/${newName}`;
                        link.download = newName;
                        link.innerText = `下载 ${oldName} → ${newName}`;
                        downloadsDiv.appendChild(link);
                    }
                });
            }

            if(data.zip_url){
                const zipLink = document.createElement("a");
                zipLink.href = data.zip_url;
                zipLink.download = "converted_files.zip";
                zipLink.innerText = "下载全部转换文件 (ZIP)";
                downloadsDiv.appendChild(document.createElement("br"));
                downloadsDiv.appendChild(zipLink);
            }
        } catch(err){
            resultDiv.innerText = "❌ 请求失败：" + err;
        }
    });

    // 清空临时文件按钮
    document.getElementById('clearTempBtn').addEventListener('click', async function(){
        const clearResult = document.getElementById('clearResult');
        clearResult.innerText = "⏳ 正在清空临时文件...";
        try {
            const res = await fetch('/clear_temp', {method:'POST'});
            const data = await res.json();
            clearResult.innerText = data.msg || "清空完成";
            
            // 清空文件选择
            selectedFiles = [];
            document.getElementById('files').value = '';
            const oldFileList = document.querySelector('.file-list');
            if (oldFileList) {
                oldFileList.remove();
            }
        } catch(err){
            clearResult.innerText = "❌ 清空失败：" + err;
        }
    });
}

// 初始化
init();
