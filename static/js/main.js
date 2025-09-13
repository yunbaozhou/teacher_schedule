// 全局变量存储课程数据
let courseData = [];
let userSelectedColors = {};

// 生成时间选项（5分钟为间隔）
function generateTimeOptions() {
    const startSelect = document.getElementById('edit-course-start-time');
    const endSelect = document.getElementById('edit-course-end-time');
    
    // 清空现有选项
    startSelect.innerHTML = '<option value="">选择开始时间</option>';
    endSelect.innerHTML = '<option value="">选择结束时间</option>';
    
    // 定义不同时段的时间范围
    const timeRanges = [
        { start: 6, end: 12, label: '上午' },   // 上午 6:00-12:00
        { start: 12, end: 18, label: '下午' },  // 下午 12:00-18:00
        { start: 18, end: 23, label: '晚自习' } // 晚自习 18:00-23:00
    ];
    
    // 为每个时段生成时间选项
    timeRanges.forEach(range => {
        // 添加时段分隔选项（仅作显示用途）
        const groupStart = document.createElement('optgroup');
        groupStart.label = range.label;
        startSelect.appendChild(groupStart);
        
        const groupEnd = document.createElement('optgroup');
        groupEnd.label = range.label;
        endSelect.appendChild(groupEnd);
        
        // 生成该时段的具体时间选项
        for (let hour = range.start; hour < range.end; hour++) {
            for (let minute = 0; minute < 60; minute += 5) {
                const timeString = `${hour.toString().padStart(2, '0')}:${minute.toString().padStart(2, '0')}`;
                const option1 = document.createElement('option');
                option1.value = timeString;
                option1.textContent = timeString;
                groupStart.appendChild(option1);
                
                const option2 = document.createElement('option');
                option2.value = timeString;
                option2.textContent = timeString;
                groupEnd.appendChild(option2);
            }
        }
    });
}

// 页面加载完成后初始化
document.addEventListener('DOMContentLoaded', function() {
    // 初始化时间选择器选项
    const startSelect = document.getElementById('edit-course-start-time');
    const endSelect = document.getElementById('edit-course-end-time');
    
    // 添加时间选项（以5分钟为间隔，从6点开始）
    for (let hour = 6; hour < 24; hour++) {
        for (let minute = 0; minute < 60; minute += 5) {
            const timeString = `${hour.toString().padStart(2, '0')}:${minute.toString().padStart(2, '0')}`;
            const option1 = new Option(timeString, timeString);
            const option2 = new Option(timeString, timeString);
            startSelect.add(option1);
            endSelect.add(option2);
        }
    }
    
    // 绑定表格单元格点击事件
    document.querySelectorAll('#course-table tbody td[data-period]').forEach(cell => {
        cell.addEventListener('click', function(e) {
            // 防止点击课程卡片时触发
            if (e.target.classList.contains('course-card') || e.target.closest('.course-card')) return;
            
            const period = parseInt(this.getAttribute('data-period'));
            const day = this.getAttribute('data-day');
            openAddModal(day, period);
        });
    });
    
    // 绑定模态框关闭事件
    document.querySelector('.close-modal').addEventListener('click', closeModal);
    document.getElementById('edit-modal').addEventListener('click', function(e) {
        if (e.target === this) {
            closeModal();
        }
    });
    
    // 绑定保存按钮事件
    document.getElementById('save-course-btn').addEventListener('click', saveCourse);
    
    // 绑定删除按钮事件
    document.getElementById('delete-course-btn').addEventListener('click', deleteCourse);
    
    // 绑定添加星期节次按钮事件
    document.getElementById('add-schedule-item').addEventListener('click', function() {
        addScheduleSelectionItem();
    });
    
    // 绑定删除星期节次选择项按钮事件
    document.addEventListener('click', function(e) {
        if (e.target.closest('.remove-schedule-item')) {
            const item = e.target.closest('.schedule-selection-item');
            const selectionList = document.getElementById('schedule-selection-list');
            if (selectionList.children.length > 1) {
                selectionList.removeChild(item);
            } else {
                alert('至少需要保留一个星期节次选择项');
            }
        }
    });
    
    // 绑定颜色选择器事件
    document.querySelectorAll('.color-picker').forEach(picker => {
        picker.addEventListener('click', function() {
            // 移除其他选择器的选中状态
            document.querySelectorAll('.color-picker').forEach(p => {
                p.classList.remove('selected-color');
            });
            
            // 添加当前选择器的选中状态
            this.classList.add('selected-color');
            
            // 设置选中的颜色值
            const color = this.getAttribute('data-color');
            document.getElementById('selected-color').value = color;
        });
    });
    
    // 绑定自定义颜色选择器事件
    document.getElementById('custom-color').addEventListener('input', function() {
        // 移除其他选择器的选中状态
        document.querySelectorAll('.color-picker').forEach(p => {
            p.classList.remove('selected-color');
        });
        
        // 获取选择的颜色并转换为RGB格式
        const hex = this.value;
        const rgb = hexToRgb(hex);
        if (rgb) {
            document.getElementById('selected-color').value = `${rgb.r},${rgb.g},${rgb.b}`;
        }
    });
    
    // 绑定右键菜单事件
    document.addEventListener('contextmenu', function(e) {
        const courseCard = e.target.closest('.course-card');
        if (courseCard) {
            e.preventDefault();
            showContextMenu(e.clientX, e.clientY, courseCard);
        }
    });
    
    // 点击其他地方隐藏右键菜单
    document.addEventListener('click', function() {
        const contextMenu = document.getElementById('context-menu');
        if (contextMenu) {
            contextMenu.style.display = 'none';
        }
    });
    
    // 绑定导出按钮事件
    document.getElementById('export-excel').addEventListener('click', exportToExcel);
    document.getElementById('export-word').addEventListener('click', exportToWord);
    document.getElementById('export-image').addEventListener('click', exportToImage);
    document.getElementById('print-schedule').addEventListener('click', printSchedule);
    
    // 绑定导入按钮事件
    document.getElementById('import-schedule').addEventListener('click', function() {
        document.getElementById('import-file').click();
    });
    
    document.getElementById('import-file').addEventListener('change', importFromExcel);
    
    // 绑定复制和粘贴按钮事件
    document.getElementById('copy-course').addEventListener('click', copyCourse);
    document.getElementById('paste-course').addEventListener('click', pasteCourse);
});

// 初始化颜色选择器
function initColorPickers() {
    // 预定义颜色选择
    document.querySelectorAll('.color-picker').forEach(picker => {
        picker.addEventListener('click', function() {
            // 移除所有选中状态
            document.querySelectorAll('.color-picker').forEach(p => {
                p.classList.remove('selected-color');
            });
            
            // 添加选中状态到当前元素
            this.classList.add('selected-color');
            
            // 保存选中的颜色
            document.getElementById('selected-color').value = this.getAttribute('data-color');
        });
    });
    
    // 自定义颜色选择器
    document.getElementById('custom-color').addEventListener('change', function() {
        // 移除所有预定义颜色的选中状态
        document.querySelectorAll('.color-picker').forEach(p => {
            p.classList.remove('selected-color');
        });
        
        // 获取选择的颜色并转换为RGB
        const color = this.value;
        const r = parseInt(color.substr(1, 2), 16);
        const g = parseInt(color.substr(3, 2), 16);
        const b = parseInt(color.substr(5, 2), 16);
        
        document.getElementById('selected-color').value = `${r},${g},${b}`;
    });
}

// 打开添加课程模态框
function openAddModal(day, period) {
    // 设置模态框标题
    document.getElementById('modal-title').textContent = '添加课程';
    
    // 设置隐藏字段
    document.getElementById('edit-mode').value = 'add';
    document.getElementById('edit-course-index').value = '';
    
    // 填充默认值
    document.getElementById('edit-course-name').value = '';
    document.getElementById('edit-course-teacher').value = '';
    document.getElementById('edit-course-location').value = '';
    document.getElementById('edit-course-notes').value = '';
    document.getElementById('edit-course-start-time').value = '';
    document.getElementById('edit-course-end-time').value = '';
    document.getElementById('selected-color').value = '';
    
    // 清空之前的星期节次选择项
    const selectionList = document.getElementById('schedule-selection-list');
    selectionList.innerHTML = '';
    
    // 添加默认选择项
    addScheduleSelectionItem(day, period);
    
    // 显示添加星期节次按钮
    document.getElementById('add-schedule-item').style.display = 'block';
    
    // 隐藏删除按钮
    document.getElementById('delete-course-btn').classList.add('hidden');
    
    // 移除所有颜色选择器的选中状态
    document.querySelectorAll('.color-picker').forEach(p => {
        p.classList.remove('selected-color');
    });
    document.getElementById('custom-color').value = '#000000';
    
    // 显示模态框
    document.getElementById('edit-modal').style.display = 'block';
}

// 打开编辑课程模态框
function openEditModal(index) {
    const course = courseData[index];
    
    // 设置模态框标题
    document.getElementById('modal-title').textContent = '编辑课程';
    
    // 设置隐藏字段
    document.getElementById('edit-mode').value = 'edit';
    document.getElementById('edit-course-index').value = index;
    
    // 填充课程数据
    document.getElementById('edit-course-name').value = course.课程名称;
    document.getElementById('edit-course-teacher').value = course.教师 || '';
    document.getElementById('edit-course-location').value = course.地点 || '';
    document.getElementById('edit-course-notes').value = course.备注 || '';
    document.getElementById('edit-course-start-time').value = course.开始时间 || '';
    document.getElementById('edit-course-end-time').value = course.结束时间 || '';
    
    // 设置颜色选择器
    const color = course.颜色;
    document.getElementById('selected-color').value = color;
    
    // 移除所有颜色选择器的选中状态
    document.querySelectorAll('.color-picker').forEach(p => {
        p.classList.remove('selected-color');
    });
    document.getElementById('custom-color').value = '#000000';
    
    // 查找匹配的颜色选择器并选中
    const colorPickers = document.querySelectorAll('.color-picker');
    let colorMatched = false;
    for (let i = 0; i < colorPickers.length; i++) {
        if (colorPickers[i].getAttribute('data-color') === color) {
            colorPickers[i].classList.add('selected-color');
            colorMatched = true;
            break;
        }
    }
    
    // 如果没有匹配的颜色选择器，设置自定义颜色
    if (!colorMatched && color) {
        const rgbValues = color.split(',');
        if (rgbValues.length === 3) {
            const r = parseInt(rgbValues[0]);
            const g = parseInt(rgbValues[1]);
            const b = parseInt(rgbValues[2]);
            const hex = rgbToHex(r, g, b);
            document.getElementById('custom-color').value = hex;
        }
    }
    
    // 清空之前的星期节次选择项
    const selectionList = document.getElementById('schedule-selection-list');
    selectionList.innerHTML = '';
    
    // 只为当前选中的课程创建一个选择项
    const newItem = document.createElement('div');
    newItem.className = 'schedule-selection-item grid grid-cols-12 gap-2 mb-2';
    newItem.innerHTML = `
        <div class="col-span-5">
            <select class="schedule-day w-full px-2 py-1 border border-gray-300 rounded-md text-sm">
                <option value="周一" ${course.星期 === '周一' ? 'selected' : ''}>星期一</option>
                <option value="周二" ${course.星期 === '周二' ? 'selected' : ''}>星期二</option>
                <option value="周三" ${course.星期 === '周三' ? 'selected' : ''}>星期三</option>
                <option value="周四" ${course.星期 === '周四' ? 'selected' : ''}>星期四</option>
                <option value="周五" ${course.星期 === '周五' ? 'selected' : ''}>星期五</option>
                <option value="周六" ${course.星期 === '周六' ? 'selected' : ''}>星期六</option>
                <option value="周日" ${course.星期 === '周日' ? 'selected' : ''}>星期日</option>
            </select>
        </div>
        <div class="col-span-5">
            <div class="periods-container border border-gray-300 rounded-md p-2 h-24 overflow-y-auto">
                <div class="grid grid-cols-2 gap-1">
                    <div><label class="inline-flex items-center"><input type="checkbox" class="period-checkbox" value="1" ${course.节次 === 1 ? 'checked' : ''} ${course.节次 !== 1 ? 'disabled' : ''}><span class="ml-1 text-xs ${course.节次 !== 1 ? 'text-gray-400' : ''}">第1节</span></label></div>
                    <div><label class="inline-flex items-center"><input type="checkbox" class="period-checkbox" value="2" ${course.节次 === 2 ? 'checked' : ''} ${course.节次 !== 2 ? 'disabled' : ''}><span class="ml-1 text-xs ${course.节次 !== 2 ? 'text-gray-400' : ''}">第2节</span></label></div>
                    <div><label class="inline-flex items-center"><input type="checkbox" class="period-checkbox" value="3" ${course.节次 === 3 ? 'checked' : ''} ${course.节次 !== 3 ? 'disabled' : ''}><span class="ml-1 text-xs ${course.节次 !== 3 ? 'text-gray-400' : ''}">第3节</span></label></div>
                    <div><label class="inline-flex items-center"><input type="checkbox" class="period-checkbox" value="4" ${course.节次 === 4 ? 'checked' : ''} ${course.节次 !== 4 ? 'disabled' : ''}><span class="ml-1 text-xs ${course.节次 !== 4 ? 'text-gray-400' : ''}">第4节</span></label></div>
                    <div><label class="inline-flex items-center"><input type="checkbox" class="period-checkbox" value="5" ${course.节次 === 5 ? 'checked' : ''} ${course.节次 !== 5 ? 'disabled' : ''}><span class="ml-1 text-xs ${course.节次 !== 5 ? 'text-gray-400' : ''}">第5节</span></label></div>
                    <div><label class="inline-flex items-center"><input type="checkbox" class="period-checkbox" value="6" ${course.节次 === 6 ? 'checked' : ''} ${course.节次 !== 6 ? 'disabled' : ''}><span class="ml-1 text-xs ${course.节次 !== 6 ? 'text-gray-400' : ''}">第6节</span></label></div>
                    <div><label class="inline-flex items-center"><input type="checkbox" class="period-checkbox" value="7" ${course.节次 === 7 ? 'checked' : ''} ${course.节次 !== 7 ? 'disabled' : ''}><span class="ml-1 text-xs ${course.节次 !== 7 ? 'text-gray-400' : ''}">第7节</span></label></div>
                    <div><label class="inline-flex items-center"><input type="checkbox" class="period-checkbox" value="8" ${course.节次 === 8 ? 'checked' : ''} ${course.节次 !== 8 ? 'disabled' : ''}><span class="ml-1 text-xs ${course.节次 !== 8 ? 'text-gray-400' : ''}">第8节</span></label></div>
                    <div><label class="inline-flex items-center"><input type="checkbox" class="period-checkbox" value="9" ${course.节次 === 9 ? 'checked' : ''} ${course.节次 !== 9 ? 'disabled' : ''}><span class="ml-1 text-xs ${course.节次 !== 9 ? 'text-gray-400' : ''}">第9节</span></label></div>
                    <div><label class="inline-flex items-center"><input type="checkbox" class="period-checkbox" value="10" ${course.节次 === 10 ? 'checked' : ''} ${course.节次 !== 10 ? 'disabled' : ''}><span class="ml-1 text-xs ${course.节次 !== 10 ? 'text-gray-400' : ''}">第10节</span></label></div>
                </div>
            </div>
        </div>
        <div class="col-span-2 flex items-center">
            <button type="button" class="remove-schedule-item text-red-500 hover:text-red-700">
                <i class="fas fa-trash"></i>
            </button>
        </div>
    `;
    
    selectionList.appendChild(newItem);
    
    // 绑定删除事件
    newItem.querySelector('.remove-schedule-item').addEventListener('click', function() {
        if (selectionList.children.length > 1) {
            selectionList.removeChild(newItem);
        } else {
            alert('至少需要保留一个星期节次选择项');
        }
    });
    
    // 隐藏添加星期节次按钮
    document.getElementById('add-schedule-item').style.display = 'none';
    
    // 显示删除按钮
    document.getElementById('delete-course-btn').classList.remove('hidden');
    
    // 显示模态框
    document.getElementById('edit-modal').style.display = 'block';
}

// 添加星期节次选择项
function addScheduleSelectionItem(day, period) {
    const selectionList = document.getElementById('schedule-selection-list');
    const newItem = document.createElement('div');
    newItem.className = 'schedule-selection-item grid grid-cols-12 gap-2 mb-2';
    newItem.innerHTML = `
        <div class="col-span-5">
            <select class="schedule-day w-full px-2 py-1 border border-gray-300 rounded-md text-sm">
                <option value="周一">星期一</option>
                <option value="周二">星期二</option>
                <option value="周三">星期三</option>
                <option value="周四">星期四</option>
                <option value="周五">星期五</option>
                <option value="周六">星期六</option>
                <option value="周日">星期日</option>
            </select>
        </div>
        <div class="col-span-5">
            <div class="periods-container border border-gray-300 rounded-md p-2 h-24 overflow-y-auto">
                <div class="grid grid-cols-2 gap-1">
                    <div><label class="inline-flex items-center"><input type="checkbox" class="period-checkbox" value="1"><span class="ml-1 text-xs">第1节</span></label></div>
                    <div><label class="inline-flex items-center"><input type="checkbox" class="period-checkbox" value="2"><span class="ml-1 text-xs">第2节</span></label></div>
                    <div><label class="inline-flex items-center"><input type="checkbox" class="period-checkbox" value="3"><span class="ml-1 text-xs">第3节</span></label></div>
                    <div><label class="inline-flex items-center"><input type="checkbox" class="period-checkbox" value="4"><span class="ml-1 text-xs">第4节</span></label></div>
                    <div><label class="inline-flex items-center"><input type="checkbox" class="period-checkbox" value="5"><span class="ml-1 text-xs">第5节</span></label></div>
                    <div><label class="inline-flex items-center"><input type="checkbox" class="period-checkbox" value="6"><span class="ml-1 text-xs">第6节</span></label></div>
                    <div><label class="inline-flex items-center"><input type="checkbox" class="period-checkbox" value="7"><span class="ml-1 text-xs">第7节</span></label></div>
                    <div><label class="inline-flex items-center"><input type="checkbox" class="period-checkbox" value="8"><span class="ml-1 text-xs">第8节</span></label></div>
                    <div><label class="inline-flex items-center"><input type="checkbox" class="period-checkbox" value="9"><span class="ml-1 text-xs">第9节</span></label></div>
                    <div><label class="inline-flex items-center"><input type="checkbox" class="period-checkbox" value="10"><span class="ml-1 text-xs">第10节</span></label></div>
                </div>
            </div>
        </div>
        <div class="col-span-2 flex items-center">
            <button type="button" class="remove-schedule-item text-red-500 hover:text-red-700">
                <i class="fas fa-trash"></i>
            </button>
        </div>
    `;
    selectionList.appendChild(newItem);
    
    // 如果提供了默认值，则设置
    if (day) {
        newItem.querySelector('.schedule-day').value = day;
    }
    
    if (period) {
        // 在复选框中选中指定的节次
        const checkbox = newItem.querySelector(`.period-checkbox[value="${period}"]`);
        if (checkbox) {
            checkbox.checked = true;
        }
    }
    
    // 绑定删除事件
    newItem.querySelector('.remove-schedule-item').addEventListener('click', function() {
        if (selectionList.children.length > 1) {
            selectionList.removeChild(newItem);
        } else {
            alert('至少需要保留一个星期节次选择项');
        }
    });
}

// 关闭模态框
function closeModal() {
    document.getElementById('edit-modal').style.display = 'none';
}

// 保存课程（添加或编辑）
function saveCourse() {
    const mode = document.getElementById('edit-mode').value;
    const courseName = document.getElementById('edit-course-name').value;
    const courseTeacher = document.getElementById('edit-course-teacher').value;
    const courseLocation = document.getElementById('edit-course-location').value;
    const courseNotes = document.getElementById('edit-course-notes').value;
    const startTime = document.getElementById('edit-course-start-time').value;
    const endTime = document.getElementById('edit-course-end-time').value;
    const selectedColor = document.getElementById('selected-color').value;
    
    if (!courseName) {
        alert('请输入课程名称');
        return;
    }
    
    // 获取所有星期节次选择项
    const scheduleItems = document.querySelectorAll('.schedule-selection-item');
    const scheduleData = [];
    
    // 检查是否有选中的节次
    let hasSelectedPeriod = false;
    
    scheduleItems.forEach(item => {
        const day = item.querySelector('.schedule-day').value;
        let selectedPeriods = [];
        
        // 获取选中的复选框（添加和编辑模式现在都使用复选框）
        const periodCheckboxes = item.querySelectorAll('.period-checkbox:checked');
        
        if (periodCheckboxes.length > 0) {
            // 使用复选框
            selectedPeriods = Array.from(periodCheckboxes).map(checkbox => parseInt(checkbox.value));
        }
        
        if (selectedPeriods.length > 0) {
            hasSelectedPeriod = true;
            scheduleData.push({
                day: day,
                periods: selectedPeriods
            });
        }
    });
    
    if (!hasSelectedPeriod) {
        alert('请至少选择一个节次');
        return;
    }
    
    // 构建时间字符串
    let courseTime = '';
    if (startTime && endTime) {
        courseTime = `${startTime}~${endTime}`;
    }
    
    if (mode === 'add') {
        // 检查冲突
        let hasConflict = false;
        const conflicts = [];
        
        scheduleData.forEach(data => {
            data.periods.forEach(period => {
                const existingCourse = courseData.find(c => 
                    c.星期 === data.day && 
                    c.节次 === period
                );
                
                if (existingCourse) {
                    hasConflict = true;
                    conflicts.push(`星期${data.day} 第${period}节：${existingCourse.课程名称}`);
                }
            });
        });
        
        if (hasConflict) {
            const confirmResult = confirm('以下时间段存在冲突：\n' + conflicts.join('\n') + '\n是否继续添加？');
            if (!confirmResult) {
                return;
            }
        }
        
        // 添加新课程
        scheduleData.forEach(data => {
            data.periods.forEach(period => {
                const newCourse = {
                    课程名称: courseName,
                    星期: data.day,
                    节次: period,
                    教师: courseTeacher,
                    地点: courseLocation,
                    备注: courseNotes,
                    开始时间: startTime,
                    结束时间: endTime,
                    颜色: selectedColor
                };
                
                courseData.push(newCourse);
            });
        });
        
        // 更新表格
        updateCourseTable();
        
        // 检查冲突
        checkConflicts();
        
        // 关闭模态框
        closeModal();
    } else {
        // 编辑模式 - 只更新当前正在编辑的课程
        const index = parseInt(document.getElementById('edit-course-index').value);
        const originalCourse = courseData[index];
        
        // 获取当前编辑的星期和节次（这些是不变的）
        const currentDay = originalCourse.星期;
        const currentPeriod = originalCourse.节次;
        
        // 删除当前正在编辑的课程
        courseData.splice(index, 1);
        
        // 创建更新后的课程信息（保持原来的星期和节次）
        const updatedCourse = {
            课程名称: courseName,
            星期: currentDay,
            节次: currentPeriod,
            教师: courseTeacher,
            地点: courseLocation,
            备注: courseNotes,
            开始时间: startTime,
            结束时间: endTime,
            颜色: selectedColor
        };
        
        // 将更新后的课程添加回数据数组
        courseData.push(updatedCourse);
        
        // 更新表格
        updateCourseTable();
        
        // 检查冲突
        checkConflicts();
        
        // 关闭模态框
        closeModal();
    }
}

// 删除课程
function deleteCourse() {
    if (confirm('确定要删除这个课程吗？')) {
        const index = parseInt(document.getElementById('edit-course-index').value);
        courseData.splice(index, 1);
        
        // 关闭模态框
        closeModal();
        
        // 更新表格显示
        updateCourseTable();
        
        // 检查冲突
        checkConflicts();
    }
}

// 检查时间冲突
function checkTimeConflict(day, period, excludeIndex = -1) {
    for (let i = 0; i < courseData.length; i++) {
        if (i === excludeIndex) continue;
        if (courseData[i].星期 === day && courseData[i].节次 === period) {
            return `课程：${courseData[i].课程名称} 教师：${courseData[i].教师}`;
        }
    }
    return null;
}

// 获取课程颜色
function getCourseColor(courseName) {
    // 预定义颜色映射
    const colorMap = {
        "语文": [255, 204, 204],
        "数学": [204, 255, 255],
        "英语": [204, 255, 204],
        "综研": [229, 229, 204],
        "趣味体育": [255, 255, 153],
        "体育": [153, 204, 255],
        "音乐": [221, 170, 221],
        "体育与健康": [153, 204, 255],
        "道法": [204, 153, 204],
        "美术": [255, 179, 136],
        "科学": [255, 229, 153],
        "劳动": [204, 255, 204]
    };
    
    // 额外的默认颜色（扩展颜色池以减少冲突）
    const defaultColors = [
        [255, 192, 203], [173, 216, 230], [144, 238, 144], [255, 182, 193],
        [221, 160, 221], [175, 238, 238], [255, 218, 185], [240, 230, 140],
        [230, 230, 250], [255, 228, 196], [255, 160, 122], [176, 224, 230],
        [255, 228, 181], [189, 183, 107], [216, 191, 216], [152, 251, 152],
        [173, 216, 230], [255, 192, 203], [244, 164, 96], [210, 180, 140],
        [255, 215, 0], [218, 112, 214], [192, 192, 192], [128, 128, 0],
        [128, 0, 128], [0, 128, 128], [0, 0, 128], [139, 0, 0],
        [0, 100, 0], [128, 0, 0]
    ];
    
    // 如果用户选择了颜色，则优先使用用户选择的颜色
    if (userSelectedColors[courseName]) {
        return userSelectedColors[courseName];
    }
    
    // 如果预定义了颜色，则使用预定义颜色
    if (colorMap[courseName]) {
        return colorMap[courseName];
    }
    
    // 基于课程名称生成稳定的哈希值，确保相同名称的课程总是获得相同的颜色
    // 使用更大的颜色池以减少不同名称课程获得相同颜色的概率
    let hash = 0;
    for (let i = 0; i < courseName.length; i++) {
        const char = courseName.charCodeAt(i);
        hash = ((hash << 5) - hash) + char;
        hash = hash & hash; // 转换为32位整数
    }
    
    // 使用扩展的颜色池
    const colorIndex = Math.abs(hash) % defaultColors.length;
    return defaultColors[colorIndex];
}

// 更新课程表显示
function updateCourseTable() {
    // 清空表格
    document.querySelectorAll('#course-table tbody td[data-period]').forEach(cell => {
        cell.innerHTML = '';
    });
    
    // 填充课程数据
    courseData.forEach((course, index) => {
        const cell = document.querySelector(`#course-table tbody td[data-period="${course.节次}"][data-day="${course.星期}"]`);
        if (cell) {
            // 创建课程卡片
            const card = document.createElement('div');
            card.className = 'course-card p-2 m-1 rounded-md';
            
            // 构建卡片内容，只显示存在的信息
            let cardContent = `<div class="font-bold text-lg">${course.课程名称}</div>`;
            
            // 只有当教师信息存在且不为"未指定"时才显示
            if (course.教师 && course.教师 !== '未指定') {
                cardContent += `<div class="text-xs">教师：${course.教师}</div>`;
            }
            
            // 只有当地点信息存在且不为"未指定"时才显示
            if (course.地点 && course.地点 !== '未指定') {
                cardContent += `<div class="text-xs">地点：${course.地点}</div>`;
            }
            
            // 只有当时间段信息存在时才显示
            if (course.开始时间 && course.结束时间) {
                cardContent += `<div class="text-xs font-medium text-blue-700">时间：${course.开始时间}~${course.结束时间}</div>`;
            }
            
            // 只有当备注信息存在时才显示
            if (course.备注) {
                cardContent += `<div class="notes-text">备注：${course.备注}</div>`;
            }
            
            card.innerHTML = cardContent;
            card.dataset.index = index;
            
            // 应用保存的颜色，如果没有则使用默认颜色策略
            let bgColor = '';
            if (course.颜色) {
                // 如果课程有保存的颜色，则使用保存的颜色
                const rgbValues = course.颜色.split(',');
                if (rgbValues.length === 3) {
                    bgColor = `rgb(${rgbValues[0]}, ${rgbValues[1]}, ${rgbValues[2]})`;
                }
            }
            
            if (!bgColor) {
                // 如果没有保存的颜色，则使用基于课程名称生成的颜色
                const color = getCourseColor(course.课程名称);
                bgColor = `rgb(${color[0]}, ${color[1]}, ${color[2]})`;
            }
            
            card.style.backgroundColor = bgColor;
            
            // 添加点击事件
            card.addEventListener('click', function(e) {
                e.stopPropagation(); // 阻止事件冒泡
                openEditModal(index);
            });
            
            cell.appendChild(card);
        }
    });
}

// 检查冲突
function checkConflicts() {
    // 创建一个映射来跟踪每个时间段的课程
    const timeMap = new Map();
    
    // 检查所有课程是否存在冲突
    const conflicts = [];
    
    courseData.forEach(course => {
        const key = `${course.星期}-${course.节次}`;
        
        if (timeMap.has(key)) {
            // 发现冲突
            const existingCourse = timeMap.get(key);
            conflicts.push(`星期${course.星期} 第${course.节次}节：${course.课程名称} 与 ${existingCourse.课程名称} 冲突`);
        } else {
            // 添加当前课程到映射中
            timeMap.set(key, course);
        }
    });
    
    // 显示冲突信息
    const conflictAlert = document.getElementById('conflict-alert');
    const conflictList = document.getElementById('conflict-list');
    
    if (conflicts.length > 0) {
        conflictList.innerHTML = '';
        conflicts.forEach(conflict => {
            const li = document.createElement('li');
            li.textContent = conflict;
            conflictList.appendChild(li);
        });
        conflictAlert.classList.remove('hidden');
    } else {
        conflictAlert.classList.add('hidden');
    }
}

// 导出到Excel
function exportToExcel() {
    if (courseData.length === 0) {
        alert('请先添加课程数据');
        return;
    }
    
    // 获取课表标题
    const timetableTitle = document.getElementById('timetable-title').value || '课程表';
    
    // 发送请求到后端生成Excel文件
    fetch('/api/export/excel', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            courses: courseData,
            title: timetableTitle
        })
    })
    .then(response => {
        if (response.ok) {
            return response.blob();
        }
        throw new Error('导出失败');
    })
    .then(blob => {
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = timetableTitle + '.xlsx';
        document.body.appendChild(a);
        a.click();
        a.remove();
        window.URL.revokeObjectURL(url);
    })
    .catch(error => {
        console.error('导出Excel时出错:', error);
        alert('导出Excel失败: ' + error.message);
    });
}

// 导出到Word
function exportToWord() {
    if (courseData.length === 0) {
        alert('请先添加课程数据');
        return;
    }
    
    // 获取课表标题
    const timetableTitle = document.getElementById('timetable-title').value || '课程表';
    
    // 发送请求到后端生成Word文件
    fetch('/api/export/word', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            courses: courseData,
            userSelectedColors: userSelectedColors,
            title: timetableTitle
        })
    })
    .then(response => {
        if (response.ok) {
            return response.blob();
        }
        throw new Error('导出失败');
    })
    .then(blob => {
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = timetableTitle + '.docx';
        document.body.appendChild(a);
        a.click();
        a.remove();
        window.URL.revokeObjectURL(url);
    })
    .catch(error => {
        console.error('导出Word时出错:', error);
        alert('导出Word失败: ' + error.message);
    });
}

// 导出到图片
function exportToImage() {
    if (courseData.length === 0) {
        alert('请先添加课程数据');
        return;
    }
    
    // 显示加载状态
    const originalText = document.getElementById('export-image').innerHTML;
    document.getElementById('export-image').innerHTML = '<i class="fas fa-spinner fa-spin mr-1"></i>导出中...';
    document.getElementById('export-image').disabled = true;
    
    // 获取课表标题
    const timetableTitle = document.getElementById('timetable-title').value || '课程表';
    
    // 克隆表格以避免影响原始表格
    const table = document.getElementById('course-table');
    const tableClone = table.cloneNode(true);
    
    // 调整克隆表格的样式以解决显示问题
    tableClone.style.fontSize = '14px';
    tableClone.style.width = '100%';
    tableClone.style.tableLayout = 'fixed'; // 固定表格布局防止内容溢出
    
    // 调整表格单元格样式实现垂直居中
    const tableCells = tableClone.querySelectorAll('td, th');
    tableCells.forEach(cell => {
        cell.style.display = 'table-cell';
        cell.style.verticalAlign = 'middle'; // 垂直居中
    });
    
    // 调整课程卡片样式，使其与页面中的样式保持一致
    const courseCards = tableClone.querySelectorAll('.course-card');
    courseCards.forEach(card => {
        // 保持与页面中相同的样式
        card.style.padding = '8px';
        card.style.margin = '2px';
        card.style.borderRadius = '6px';
        card.style.fontSize = '12px';
        card.style.cursor = 'pointer';
        card.style.minHeight = '60px';
        card.style.display = 'flex';
        card.style.flexDirection = 'column';
        card.style.justifyContent = 'flex-start'; // 与页面中保持一致的对齐方式
        
        // 重置文本对齐方式，让内容自然对齐
        const children = card.children;
        for (let i = 0; i < children.length; i++) {
            const child = children[i];
            // 课程名称保持左对齐（与页面中保持一致）
            if (child.classList.contains('font-bold')) {
                child.style.textAlign = 'left';
            } else {
                child.style.textAlign = 'left';
            }
            child.style.width = '100%';
        }
    });
    
    // 创建一个临时容器来放置克隆的表格
    const tempContainer = document.createElement('div');
    tempContainer.style.padding = '20px';
    tempContainer.style.backgroundColor = '#f8fafc';
    tempContainer.style.width = table.offsetWidth + 'px'; // 设置容器宽度与原表格一致
    tempContainer.appendChild(tableClone);
    
    // 添加到页面中但隐藏起来
    tempContainer.style.position = 'absolute';
    tempContainer.style.left = '-9999px';
    tempContainer.style.top = '-9999px';
    document.body.appendChild(tempContainer);
    
    // 使用html2canvas将表格转换为图片
    html2canvas(tempContainer, {
        scale: 2, // 提高图片质量
        useCORS: true,
        backgroundColor: '#f8fafc',
        scrollY: -window.scrollY,
        windowHeight: tempContainer.scrollHeight + 100, // 增加额外高度确保内容完整显示
        width: tempContainer.scrollWidth, // 设置画布宽度
        height: tempContainer.scrollHeight // 设置画布高度
    }).then(canvas => {
        // 恢复按钮状态
        document.getElementById('export-image').innerHTML = originalText;
        document.getElementById('export-image').disabled = false;
        
        // 从页面中移除临时容器
        document.body.removeChild(tempContainer);
        
        // 将canvas转换为图片并下载
        const link = document.createElement('a');
        link.download = timetableTitle + '.png';
        link.href = canvas.toDataURL('image/png');
        link.click();
    }).catch(error => {
        // 恢复按钮状态
        document.getElementById('export-image').innerHTML = originalText;
        document.getElementById('export-image').disabled = false;
        
        // 从页面中移除临时容器
        document.body.removeChild(tempContainer);
        
        console.error('导出图片时出错:', error);
        alert('导出图片失败: ' + error.message);
    });
}

// 打印课程表
function printSchedule() {
    // 获取课表标题
    const timetableTitle = document.getElementById('timetable-title').value || '课程表';
    
    // 创建打印窗口
    const printWindow = window.open('', '_blank');
    
    // 获取表格HTML并进行处理
    const table = document.getElementById('course-table').cloneNode(true);
    
    // 移除事件监听器相关的属性
    table.querySelectorAll('*').forEach(element => {
        element.removeAttribute('data-period');
        element.removeAttribute('data-day');
        element.removeAttribute('id');
    });
    
    // 复制课程卡片样式
    document.querySelectorAll('.course-card').forEach(card => {
        const cloneCard = card.cloneNode(true);
        const index = card.dataset.index;
        const originalCard = table.querySelector(`.course-card[data-index="${index}"]`);
        if (originalCard) {
            originalCard.outerHTML = cloneCard.outerHTML;
        }
    });
    
    printWindow.document.write(`
        <html>
            <head>
                <title>${timetableTitle}</title>
                <style>
                    body {
                        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
                        margin: 20px;
                    }
                    table {
                        width: 100%;
                        border-collapse: collapse;
                        margin-top: 20px;
                    }
                    th, td {
                        border: 1px solid #ddd;
                        padding: 8px;
                        text-align: center;
                        vertical-align: top;
                    }
                    thead th {
                        background-color: #f3f4f6;
                        font-weight: bold;
                        padding: 8px;
                    }
                    .time-period {
                        background-color: #e5e7eb;
                        font-weight: bold;
                        text-align: center;
                        padding: 8px;
                    }
                    .course-card {
                        font-size: 12px;
                        text-align: left;
                        padding: 8px;
                        margin: 2px;
                        border-radius: 6px;
                        color: #000;
                    }
                    .notes-text {
                        font-size: 10px;
                        opacity: 0.8;
                        margin-top: 4px;
                    }
                    h1 {
                        text-align: center;
                        font-size: 24px;
                        margin-bottom: 20px;
                    }
                    .table-header {
                        background-color: #f3f4f6;
                        font-weight: bold;
                    }
                </style>
            </head>
            <body>
                <h1>${timetableTitle}</h1>
                ${table.outerHTML}
                <script>
                    window.onload = function() {
                        window.print();
                        window.onafterprint = function() {
                            window.close();
                        }
                    }
                <\/script>
            </body>
        </html>
    `);
    
    printWindow.document.close();
}

// 从Excel导入课程表数据
function importFromExcel(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    // 检查文件类型
    if (!file.name.endsWith('.xlsx')) {
        alert('请选择一个Excel文件 (.xlsx)');
        event.target.value = ''; // 清空文件选择
        return;
    }
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            
            // 获取第一个工作表
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            // 将工作表转换为JSON
            const jsonData = XLSX.utils.sheet_to_json(worksheet, {header: 1});
            
            // 解析课程数据
            parseImportedData(jsonData);
            
            // 清空文件选择
            event.target.value = '';
        } catch (error) {
            console.error('导入Excel文件时出错:', error);
            alert('导入失败: ' + error.message);
            event.target.value = ''; // 清空文件选择
        }
    };
    
    reader.onerror = function() {
        alert('读取文件时发生错误');
        event.target.value = ''; // 清空文件选择
    };
    
    reader.readAsArrayBuffer(file);
}

// 解析导入的数据
function parseImportedData(data) {
    if (!data || data.length < 3) {
        alert('Excel文件格式不正确');
        return;
    }
    
    // 清空现有课程数据
    courseData = [];
    
    // 检查第一行是否包含课表标题
    if (data[0] && data[0][0]) {
        const title = data[0][0].toString().trim();
        if (title && title !== "课表名称") {
            document.getElementById('timetable-title').value = title;
        }
    }
    
    // 定义星期映射
    const dayMap = {
        '星期一': '周一',
        '星期二': '周二',
        '星期三': '周三',
        '星期四': '周四',
        '星期五': '周五',
        '星期六': '周六',
        '星期日': '周日'
    };
    
    // 从第三行开始处理数据（前两行是标题）
    for (let row = 2; row < data.length; row++) {
        const rowData = data[row];
        if (!rowData || rowData.length === 0) continue;
        
        // 第一列是节次
        let period = parseInt(rowData[0]);
        // 如果直接解析失败，尝试解析"第N节"格式
        if (isNaN(period) && typeof rowData[0] === 'string') {
            const periodMatch = rowData[0].match(/第(\d+)节/);
            if (periodMatch && periodMatch[1]) {
                period = parseInt(periodMatch[1]);
            }
        }
        
        if (isNaN(period)) continue;
        
        // 处理各天的课程数据
        for (let col = 1; col <= 7; col++) {
            const cellData = rowData[col];
            if (!cellData) continue;
            
            // 根据列索引确定星期
            const days = ['周一', '周二', '周三', '周四', '周五', '周六', '周日'];
            const day = days[col - 1];
            
            // 解析单元格中的课程信息
            parseCourseCell(cellData, day, period);
        }
    }
    
    // 更新表格显示
    updateCourseTable();
    
    // 检查冲突
    checkConflicts();

}

// 解析单元格中的课程信息
function parseCourseCell(cellData, day, period) {
    // 将单元格数据转换为字符串
    const data = cellData.toString().trim();
    if (!data) return;
    
    // 分割课程信息（按换行符）
    const lines = data.split(/\r?\n/);
    if (lines.length === 0) return;
    
    // 第一行是课程名称
    const courseName = lines[0].trim();
    if (!courseName) return;
    
    // 解析其他信息
    let teacher = '';
    let location = '';
    let notes = '';
    let startTime = '';
    let endTime = '';
    
    // 处理其他行的信息
    for (let i = 1; i < lines.length; i++) {
        const line = lines[i].trim();
        if (line.startsWith('教师：')) {
            teacher = line.substring(3).trim();
        } else if (line.startsWith('地点：')) {
            location = line.substring(3).trim();
        } else if (line.startsWith('备注：')) {
            notes = line.substring(3).trim();
        } else if (line.startsWith('时间：')) {
            const timeStr = line.substring(3).trim();
            const timeParts = timeStr.split('~');
            if (timeParts.length === 2) {
                startTime = timeParts[0].trim();
                endTime = timeParts[1].trim();
            }
        }
    }
    
    // 创建课程对象
    const course = {
        课程名称: courseName,
        星期: day,
        节次: period,
        教师: teacher || '未指定',
        地点: location || '未指定',
        备注: notes,
        开始时间: startTime,
        结束时间: endTime
    };
    
    // 添加到课程数据中
    courseData.push(course);
}

// 辅助函数：将十六进制颜色转换为RGB对象
function hexToRgb(hex) {
    const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
    return result ? {
        r: parseInt(result[1], 16),
        g: parseInt(result[2], 16),
        b: parseInt(result[3], 16)
    } : null;
}

// 辅助函数：将RGB值转换为十六进制颜色
function rgbToHex(r, g, b) {
    return "#" + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1);
}