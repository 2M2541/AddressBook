// 页面加载时自动加载联系人列表
window.onload = loadContacts;

// -------------------------- 基础功能：增删查 --------------------------
// 1. 加载所有联系人
function loadContacts() {
    fetch('http://localhost:5000/api/contacts')
        .then(res => res.json())
        .then(contacts => {
            const container = document.getElementById('contactsContainer');
            container.innerHTML = '';

            if (contacts.length === 0) {
                container.innerHTML = '<p style="color:#666;text-align:center;padding:20px;">暂无联系人</p>';
                return;
            }

            contacts.forEach(contact => {
                const card = document.createElement('div');
                card.className = 'contact-card';
                card.innerHTML = `
                    <div class="contact-info">
                        <h4>${contact.name}</h4>
                        <p>电话：${contact.phone}</p>
                    </div>
                    <button class="delete-btn" onclick="deleteContact(${contact.id})">删除</button>
                `;
                container.appendChild(card);
            });
        })
        .catch(err => {
            console.error('加载失败：', err);
            alert('加载联系人失败，请检查后端服务是否启动！');
        });
}

// 2. 新增联系人
function addContact() {
    const name = document.getElementById('nameInput').value.trim();
    const phone = document.getElementById('phoneInput').value.trim();

    if (!name || !phone) {
        alert('姓名和电话不能为空！');
        return;
    }

    fetch('http://localhost:5000/api/contacts', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name, phone })
    })
    .then(res => res.json())
    .then(data => {
        if (data.message) {
            alert(data.message);
            // 清空输入框
            document.getElementById('nameInput').value = '';
            document.getElementById('phoneInput').value = '';
            loadContacts(); // 刷新列表
        } else {
            alert('添加失败：' + data.error);
        }
    })
    .catch(err => {
        console.error('新增失败：', err);
        alert('新增失败，请检查网络！');
    });
}

// 3. 删除联系人
function deleteContact(contactId) {
    if (!confirm('确定删除该联系人？删除后无法恢复！')) return;

    fetch(`http://localhost:5000/api/contacts/${contactId}`, { method: 'DELETE' })
        .then(res => res.json())
        .then(data => {
            if (data.message) {
                alert(data.message);
                loadContacts(); // 刷新列表
            } else {
                alert('删除失败：' + data.error);
            }
        })
        .catch(err => {
            console.error('删除失败：', err);
            alert('删除失败，请检查网络！');
        });
}

// -------------------------- Excel导出功能（优化版） --------------------------
function exportContacts() {
    // 改用fetch接收二进制流，避免window.open导致的文件截断
    fetch('http://localhost:5000/api/contacts/export')
        .then(res => {
            if (!res.ok) throw new Error('导出请求失败');
            return res.blob(); // 接收二进制文件流
        })
        .then(blob => {
            // 创建下载链接并触发下载
            const url = window.URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = url;
            link.download = '通讯录导出.xlsx'; // 强制指定文件名和扩展名
            document.body.appendChild(link);
            link.click();
            // 清理临时资源，避免内存泄漏
            setTimeout(() => {
                document.body.removeChild(link);
                window.URL.revokeObjectURL(url);
            }, 100);
        })
        .catch(err => {
            console.error('导出失败：', err);
            alert('导出失败，请检查后端服务是否正常运行！');
        });
}

// -------------------------- Excel导入功能 --------------------------
function importContacts(fileInput) {
    const file = fileInput.files[0];
    if (!file) return;
