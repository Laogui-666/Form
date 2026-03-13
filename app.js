// ============================================
// 盼达文旅 - 申根填表助手 (完全重写版)
// ============================================

// 全局变量
let currentStep = 1;
const totalSteps = 6;
let companions = [];

// ============================================
// 初始化
// ============================================
document.addEventListener('DOMContentLoaded', function() {
    initEventListeners();
    updateProgress();
    updateStepIndicators();
});

// 初始化所有事件监听
function initEventListeners() {
    // 出生日期变化 - 检查是否未成年（监听change和input事件）
    const birthDateInput = document.getElementById('birthDate');
    if (birthDateInput) {
        birthDateInput.addEventListener('change', checkMinor);
        birthDateInput.addEventListener('input', checkMinor); // 实时检测年龄变化
        // 页面加载时也检查一次
        if (birthDateInput.value) {
            checkMinor();
        }
    }
    
    // 婚姻状况变化
    const maritalStatus = document.getElementById('maritalStatus');
    if (maritalStatus) {
        maritalStatus.addEventListener('change', function() {
            toggleField('maritalStatusOtherGroup', this.value === 'other');
        });
    }
    
    // 护照种类变化
    const passportType = document.getElementById('passportType');
    if (passportType) {
        passportType.addEventListener('change', function() {
            toggleField('passportTypeOtherGroup', this.value === 'other');
        });
    }
    
    // 居住在其他国家
    const residenceAbroadRadios = document.getElementsByName('residenceAbroad');
    residenceAbroadRadios.forEach(radio => {
        radio.addEventListener('change', function() {
            const show = this.value === 'yes';
            toggleField('residencePermitGroup', show);
            toggleField('residencePermitExpiryGroup', show);
        });
    });
    
    // 职业变化
    const occupation = document.getElementById('occupation');
    if (occupation) {
        occupation.addEventListener('change', function() {
            const show = this.value === 'other';
            const employerShow = ['employed', 'self-employed', 'student'].includes(this.value);
            
            toggleField('occupationOtherGroup', show);
            toggleField('employerSection', employerShow);
            
            // 更新标签
            const label = document.getElementById('employerLabel');
            const nameLabel = document.getElementById('employerNameLabel');
            if (label) label.textContent = this.value === 'student' ? '学校信息' : '工作单位信息';
            if (nameLabel) {
                nameLabel.innerHTML = (this.value === 'student' ? '学校名称' : '工作单位名称') + ' <span class="required">*</span>';
            }
            
            // 显示职位字段
            toggleField('currentPositionGroup', this.value === 'employed');
        });
    }
    
    // 行程日期变化 - 计算天数
    const arrivalDate = document.getElementById('arrivalDate');
    const departureDate = document.getElementById('departureDate');
    if (arrivalDate) arrivalDate.addEventListener('change', calculateStayDuration);
    if (departureDate) departureDate.addEventListener('change', calculateStayDuration);
    
    // 旅程目的变化
    const tripPurposeCheckboxes = document.querySelectorAll('input[name="tripPurpose"]');
    tripPurposeCheckboxes.forEach(cb => {
        cb.addEventListener('change', function() {
            const otherChecked = Array.from(tripPurposeCheckboxes).some(c => c.value === 'other' && c.checked);
            toggleField('tripPurposeOtherGroup', otherChecked);
        });
    });
    
    // 之前申根签证
    const prevSchengenVisaRadios = document.getElementsByName('prevSchengenVisa');
    prevSchengenVisaRadios.forEach(radio => {
        radio.addEventListener('change', function() {
            toggleField('prevVisaInfoGroup', this.value === 'yes');
        });
    });
    
    // 指纹记录
    const fingerprintsRadios = document.getElementsByName('fingerprints');
    fingerprintsRadios.forEach(radio => {
        radio.addEventListener('change', function() {
            toggleField('fingerprintsDateGroup', this.value === 'yes');
        });
    });
    
    // 邀请人类型
    const hasInviterRadios = document.getElementsByName('hasInviter');
    hasInviterRadios.forEach(radio => {
        radio.addEventListener('change', function() {
            toggleField('hotelSection', this.value === 'no');
            toggleField('personalInviterSection', this.value === 'personal');
            toggleField('orgInviterSection', this.value === 'organization');
        });
    });
    
    // 费用来源
    const fundingSourceRadios = document.getElementsByName('fundingSource');
    fundingSourceRadios.forEach(radio => {
        radio.addEventListener('change', function() {
            toggleField('applicantFundingSection', this.value === 'applicant');
            toggleField('sponsorFundingSection', this.value === 'sponsor');
            
            // 如果选择"本人"，默认选中现金和信用卡
            if (this.value === 'applicant') {
                const cashCb = document.getElementById('applicantCash');
                const creditCardCb = document.getElementById('applicantCreditCard');
                if (cashCb) cashCb.checked = true;
                if (creditCardCb) creditCardCb.checked = true;
            }
        });
    });
    
    // 初始化时：如果默认选择"本人"，默认选中现金和信用卡
    const defaultFunding = document.querySelector('input[name="fundingSource"]:checked');
    if (defaultFunding && defaultFunding.value === 'applicant') {
        const cashCb = document.getElementById('applicantCash');
        const creditCardCb = document.getElementById('applicantCreditCard');
        if (cashCb) cashCb.checked = true;
        if (creditCardCb) creditCardCb.checked = true;
    }
    
    // 赞助人类型
    const sponsorTypeRadios = document.getElementsByName('sponsorType');
    sponsorTypeRadios.forEach(radio => {
        radio.addEventListener('change', function() {
            toggleField('otherSponsorNameGroup', this.value === 'other');
        });
    });
    
    // 本人支付方式
    const applicantMeansCheckboxes = document.querySelectorAll('input[name="applicantMeans"]');
    applicantMeansCheckboxes.forEach(cb => {
        cb.addEventListener('change', function() {
            const otherChecked = Array.from(applicantMeansCheckboxes).some(c => c.value === 'other' && c.checked);
            toggleField('applicantMeansOtherGroup', otherChecked);
        });
    });
    
    // 赞助人支付方式
    const sponsorMeansCheckboxes = document.querySelectorAll('input[name="sponsorMeans"]');
    sponsorMeansCheckboxes.forEach(cb => {
        cb.addEventListener('change', function() {
            const otherChecked = Array.from(sponsorMeansCheckboxes).some(c => c.value === 'other' && c.checked);
            toggleField('sponsorMeansOtherGroup', otherChecked);
        });
    });
}

// ============================================
// 辅助函数
// ============================================

// 显示/隐藏字段
function toggleField(fieldId, show) {
    const field = document.getElementById(fieldId);
    if (field) {
        field.style.display = show ? 'block' : 'none';
    }
}

// 检查是否未成年（小于18岁）
function checkMinor() {
    const birthDateInput = document.getElementById('birthDate');
    const guardianSection = document.getElementById('guardianSection');
    
    if (!birthDateInput || !birthDateInput.value || !guardianSection) return;
    
    const birthDate = new Date(birthDateInput.value);
    const today = new Date();
    
    let age = today.getFullYear() - birthDate.getFullYear();
    const monthDiff = today.getMonth() - birthDate.getMonth();
    
    if (monthDiff < 0 || (monthDiff === 0 && today.getDate() < birthDate.getDate())) {
        age--;
    }
    
    // 未成年显示监护人信息
    guardianSection.style.display = age < 18 ? 'block' : 'none';
    
    // 同时验证监护人字段（如果显示则必填）
    setGuardianRequired(age < 18);
}

// 设置监护人字段是否必填
function setGuardianRequired(required) {
    const guardianFields = ['guardian1Name', 'guardian1Nationality', 'guardian1Phone', 'guardian1Address'];
    guardianFields.forEach(id => {
        const field = document.getElementById(id);
        if (field) {
            field.required = required;
        }
    });
}

// 计算停留天数
function calculateStayDuration() {
    const arrival = document.getElementById('arrivalDate');
    const departure = document.getElementById('departureDate');
    const duration = document.getElementById('stayDuration');
    
    if (!arrival || !departure || !duration) return;
    
    if (arrival.value && departure.value) {
        const arrivalDate = new Date(arrival.value);
        const departureDate = new Date(departure.value);
        const diffTime = departureDate - arrivalDate;
        const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;
        
        duration.value = diffDays > 0 ? diffDays + ' 天' : '';
    }
}

// ============================================
// 步骤导航
// ============================================

// 下一步
function nextStep() {
    // 验证当前步骤
    if (!validateCurrentStep()) {
        return;
    }
    
    if (currentStep < totalSteps) {
        currentStep++;
        showStep(currentStep);
    }
}

// 上一步
function prevStep() {
    if (currentStep > 1) {
        currentStep--;
        showStep(currentStep);
    }
}

// 显示指定步骤
function showStep(step) {
    // 隐藏所有步骤
    document.querySelectorAll('.form-step').forEach(s => {
        s.classList.remove('active');
    });
    
    // 显示当前步骤
    const currentSection = document.getElementById('step' + step);
    if (currentSection) {
        currentSection.classList.add('active');
    }
    
    // 更新进度
    updateProgress();
    updateStepIndicators();
    
    // 如果是最后一步，生成预览
    if (step === 6) {
        generatePreview();
    }
    
    // 滚动到顶部
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

// 更新进度条
function updateProgress() {
    const progressFill = document.getElementById('progressFill');
    if (progressFill) {
        const percentage = ((currentStep - 1) / (totalSteps - 1)) * 100;
        progressFill.style.width = percentage + '%';
    }
}

// 更新步骤指示器
function updateStepIndicators() {
    document.querySelectorAll('.step').forEach((step, index) => {
        const stepNum = index + 1;
        step.classList.remove('active', 'completed');
        
        if (stepNum === currentStep) {
            step.classList.add('active');
        } else if (stepNum < currentStep) {
            step.classList.add('completed');
        }
    });
}

// 验证当前步骤
function validateCurrentStep() {
    const currentSection = document.getElementById('step' + currentStep);
    if (!currentSection) return false;
    
    // 获取当前步骤下的所有必填字段
    const requiredFields = currentSection.querySelectorAll('[required]');
    let isValid = true;
    let firstInvalid = null;
    
    requiredFields.forEach(field => {
        // 跳过隐藏字段（如果不需要填写）
        const parentSection = field.closest('.conditional-section');
        if (parentSection && parentSection.style.display === 'none') {
            return;
        }
        
        if (!field.value.trim()) {
            isValid = false;
            field.classList.add('error');
            
            if (!firstInvalid) {
                firstInvalid = field;
            }
        } else {
            field.classList.remove('error');
        }
    });
    
    // 如果有guardian section，检查是否需要填写
    if (currentStep === 1) {
        const guardianSection = document.getElementById('guardianSection');
        if (guardianSection && guardianSection.style.display !== 'none') {
            const guardianRequired = ['guardian1Name', 'guardian1Nationality', 'guardian1Phone', 'guardian1Address'];
            guardianRequired.forEach(id => {
                const field = document.getElementById(id);
                if (field && !field.value.trim()) {
                    isValid = false;
                    field.classList.add('error');
                    if (!firstInvalid) firstInvalid = field;
                }
            });
        }
    }
    
    // 如果有无效字段，滚动到第一个
    if (firstInvalid) {
        firstInvalid.scrollIntoView({ behavior: 'smooth', block: 'center' });
        firstInvalid.focus();
    }
    
    return isValid;
}

// ============================================
// 多选下拉菜单
// ============================================

// 切换目的地选择器
function toggleDestinationDropdown() {
    const dropdown = document.getElementById('destinationDropdown');
    if (dropdown) {
        dropdown.classList.toggle('show');
    }
}

// 过滤目的地
function filterDestinations() {
    const search = document.getElementById('destinationSearch').value.toLowerCase();
    const options = document.querySelectorAll('#destinationOptions .dropdown-option');
    
    options.forEach(opt => {
        const text = opt.textContent.toLowerCase();
        opt.style.display = text.includes(search) ? 'flex' : 'none';
    });
}

// 更新目的地选择
function updateDestination() {
    const checkboxes = document.querySelectorAll('#destinationOptions input[type="checkbox"]:checked');
    const selected = Array.from(checkboxes).map(cb => cb.value);
    
    const input = document.getElementById('destinationInput');
    const tagsContainer = document.getElementById('destinationTags');
    
    if (input) {
        input.value = selected.join(', ');
    }
    
    // 更新标签
    if (tagsContainer) {
        tagsContainer.innerHTML = selected.map(country => 
            `<span class="tag">${country} <span class="tag-remove" onclick="removeDestination('${country}')">×</span></span>`
        ).join('');
    }
}

// 移除目的地
function removeDestination(country) {
    const checkbox = document.querySelector(`#destinationOptions input[value="${country}"]`);
    if (checkbox) {
        checkbox.checked = false;
        updateDestination();
    }
}

// 点击外部关闭下拉菜单
document.addEventListener('click', function(e) {
    const wrapper = document.querySelector('.multi-select-wrapper');
    const dropdown = document.getElementById('destinationDropdown');
    
    if (wrapper && dropdown && !wrapper.contains(e.target)) {
        dropdown.classList.remove('show');
    }
});

// ============================================
// 同步签证申请国
// ============================================
function syncVisaCountry() {
    // 签证申请国变化时可以执行的逻辑
    // 目前暂无特殊处理需求
}

// ============================================
// 同行人管理
// ============================================

// 切换同行人 section
function toggleCompanionSection(value) {
    toggleField('companionSection', value === 'yes');
}

// 添加同行人
function addCompanion() {
    const companionList = document.getElementById('companionList');
    const index = companions.length;
    
    companions.push({ name: '', relationship: '', passportNumber: '' });
    
    const html = `
        <div class="companion-item" id="companion${index}">
            <div class="companion-header">
                <span>同行人 ${index + 1}</span>
                <button type="button" class="btn-remove" onclick="removeCompanion(${index})">删除</button>
            </div>
            <div class="form-grid">
                <div class="form-group">
                    <label>姓名</label>
                    <input type="text" name="companionName${index}" placeholder="请输入姓名">
                </div>
                <div class="form-group">
                    <label>与申请人关系</label>
                    <input type="text" name="companionRelation${index}" placeholder="如：配偶、子女、同事">
                </div>
                <div class="form-group">
                    <label>护照号码</label>
                    <input type="text" name="companionPassport${index}" placeholder="请输入护照号码">
                </div>
            </div>
        </div>
    `;
    
    if (companionList) {
        companionList.insertAdjacentHTML('beforeend', html);
    }
}

// 移除同行人
function removeCompanion(index) {
    const item = document.getElementById('companion' + index);
    if (item) {
        item.remove();
        companions.splice(index, 1);
    }
}

// ============================================
// 预览生成
// ============================================

function generatePreview() {
    const previewContent = document.getElementById('previewContent');
    if (!previewContent) return;
    
    // 收集所有表单数据
    const formData = collectFormData();
    
    // 生成预览HTML
    let html = '';
    
    // 个人信息
    html += `<div class="preview-section">
        <h4>01 个人信息</h4>
        <div class="preview-grid">
            <div><strong>姓：</strong>${formData.surname || '-'}</div>
            <div><strong>名：</strong>${formData.givenName || '-'}</div>
            <div><strong>出生日期：</strong>${formatDate(formData.birthDate)}</div>
            <div><strong>出生地：</strong>${formData.birthPlace || '-'}</div>
            <div><strong>现国籍：</strong>${formData.nationality || '-'}</div>
            <div><strong>性别：</strong>${formData.gender === 'male' ? '男' : '女'}</div>
            <div><strong>婚姻状况：</strong>${getMaritalStatusText(formData.maritalStatus)}</div>
            <div><strong>身份证号：</strong>${formData.idCard || '-'}</div>
        </div>
    </div>`;
    
    // 监护人信息（如果未成年）
    if (formData.guardian1Name) {
        html += `<div class="preview-section">
            <h4>01.1 监护人信息</h4>
            <div class="preview-grid">
                <div><strong>监护人1姓名：</strong>${formData.guardian1Name || '-'}</div>
                <div><strong>监护人1国籍：</strong>${formData.guardian1Nationality || '-'}</div>
                <div><strong>监护人1电话：</strong>${formData.guardian1Phone || '-'}</div>
                <div><strong>监护人1邮箱：</strong>${formData.guardian1Email || '-'}</div>
                <div><strong>监护人1地址：</strong>${formData.guardian1Address || '-'}</div>
            </div>
        </div>`;
    }
    
    // 证件信息
    html += `<div class="preview-section">
        <h4>02 证件与职业信息</h4>
        <div class="preview-grid">
            <div><strong>护照种类：</strong>${getPassportTypeText(formData.passportType)}</div>
            <div><strong>护照号码：</strong>${formData.passportNumber || '-'}</div>
            <div><strong>护照签发日期：</strong>${formatDate(formData.passportIssueDate)}</div>
            <div><strong>有效期至：</strong>${formatDate(formData.passportExpiry)}}</div>
            <div><strong>签发机关：</strong>${formData.passportIssuer || '-'}</div>
            <div><strong>现职业：</strong>${getOccupationText(formData.occupation)}</div>
            <div><strong>住址：</strong>${formData.address || '-'}</div>
            <div><strong>邮箱：</strong>${formData.email || '-'}</div>
        </div>
    </div>`;
    
    // 行程信息
    html += `<div class="preview-section">
        <h4>03 行程信息</h4>
        <div class="preview-grid">
            <div><strong>签证申请国：</strong>${formData.visaApplicationCountry || '-'}</div>
            <div><strong>申根目的地：</strong>${formData.schengenDestinations || '-'}</div>
            <div><strong>首入国：</strong>${formData.firstEntry || '-'}</div>
            <div><strong>入境次数：</strong>${getEntryTypeText(formData.entryType)}</div>
            <div><strong>入境日期：</strong>${formatDate(formData.arrivalDate)}</div>
            <div><strong>离境日期：</strong>${formatDate(formData.departureDate)}</div>
            <div><strong>停留天数：</strong>${formData.stayDuration || '-'}</div>
            <div><strong>主要目的：</strong>${getTripPurposeText(formData.tripPurpose)}</div>
        </div>
    </div>`;
    
    // 邀请与住宿
    html += `<div class="preview-section">
        <h4>04 邀请与住宿信息</h4>
        <div class="preview-grid">
            <div><strong>邀请类型：</strong>${getInviterTypeText(formData.hasInviter)}</div>
    `;
    
    if (formData.hasInviter === 'no') {
        html += `<div><strong>酒店名称：</strong>${formData.hotelName || '-'}</div>
            <div><strong>酒店地址：</strong>${formData.hotelAddress || '-'}</div>
            <div><strong>酒店电话：</strong>${formData.hotelPhone || '-'}</div>`;
    } else if (formData.hasInviter === 'personal') {
        html += `<div><strong>邀请人姓名：</strong>${formData.inviterName || '-'}</div>
            <div><strong>邀请人地址：</strong>${formData.inviterAddress || '-'}</div>
            <div><strong>邀请人电话：</strong>${formData.inviterPhone || '-'}</div>`;
    } else if (formData.hasInviter === 'organization') {
        html += `<div><strong>公司名称：</strong>${formData.orgName || '-'}</div>
            <div><strong>公司地址：</strong>${formData.orgAddress || '-'}</div>
            <div><strong>联系人：</strong>${formData.orgContactName || '-'}</div>`;
    }
    
    html += `</div></div>`;
    
    // 费用信息
    html += `<div class="preview-section">
        <h4>05 费用与出资信息</h4>
        <div class="preview-grid">
            <div><strong>费用来源：</strong>${formData.fundingSource === 'applicant' ? '本人支付' : '赞助人支付'}</div>
        </div>
    </div>`;
    
    previewContent.innerHTML = html;
}

// 收集表单数据
function collectFormData() {
    const data = {};
    
    // 基本信息
    const fields = ['surname', 'birthSurname', 'givenName', 'birthDate', 'birthPlace', 
        'nationality', 'birthCountry', 'idCard', 'maritalStatus', 'maritalStatusOther',
        'passportType', 'passportTypeOther', 'passportNumber', 'passportIssueDate', 
        'passportExpiry', 'passportIssuer', 'address', 'email', 'occupation', 'occupationOther',
        'employerName', 'employerAddress', 'employerPhone', 'currentPosition',
        'visaApplicationCountry', 'firstEntry', 'entryType', 'arrivalDate', 'departureDate',
        'stayDuration', 'tripPurposeOther', 'hotelName', 'hotelAddress', 'hotelPhone', 'hotelEmail',
        'inviterName', 'inviterAddress', 'inviterPhone', 'inviterEmail',
        'orgName', 'orgAddress', 'orgPhone', 'orgContactName', 'orgContactPhone', 'orgContactEmail',
        'applicantMeansOther', 'otherSponsorName', 'sponsorMeansOther',
        'guardian1Name', 'guardian1Nationality', 'guardian1Phone', 'guardian1Email', 'guardian1Address',
        'guardian2Name', 'guardian2Nationality', 'guardian2Phone', 'guardian2Email', 'guardian2Address'];
    
    fields.forEach(field => {
        const el = document.getElementById(field);
        if (el) data[field] = el.value;
    });
    
    // 单选按钮
    const radioFields = ['gender', 'maritalStatus', 'residenceAbroad', 'occupation', 
        'hasCompanion', 'prevSchengenVisa', 'fingerprints', 'hasInviter', 'fundingSource', 'sponsorType'];
    
    radioFields.forEach(name => {
        const checked = document.querySelector(`input[name="${name}"]:checked`);
        if (checked) data[name] = checked.value;
    });
    
    // 多选目的地
    const destCheckboxes = document.querySelectorAll('#destinationOptions input[type="checkbox"]:checked');
    data.schengenDestinations = Array.from(destCheckboxes).map(cb => cb.value).join(', ');
    
    // 旅程目的
    const purposeCheckboxes = document.querySelectorAll('input[name="tripPurpose"]:checked');
    data.tripPurpose = Array.from(purposeCheckboxes).map(cb => cb.value);
    
    return data;
}

// ============================================
// 格式化函数
// ============================================

function formatDate(dateStr) {
    if (!dateStr) return '-';
    const date = new Date(dateStr);
    return date.toLocaleDateString('zh-CN');
}

// 获取当前日期
function getCurrentDate() {
    const today = new Date();
    const year = today.getFullYear();
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const day = String(today.getDate()).padStart(2, '0');
    return `${year}年${month}月${day}日`;
}

function getMaritalStatusText(value) {
    const map = { 'single': '未婚', 'married': '已婚', 'separated': '分居', 
        'divorced': '离异', 'widowed': '丧偶', 'other': '其他' };
    return map[value] || value;
}

function getPassportTypeText(value) {
    const map = { 'ordinary': '普通护照', 'diplomatic': '外交护照', 
        'service': '公务护照', 'official': '因公护照', 'special': '特殊护照', 'other': '其他' };
    return map[value] || value;
}

function getOccupationText(value) {
    const map = { 'student': '学生', 'employed': '在职', 'self-employed': '自雇', 
        'retired': '退休', 'unemployed': '无业', 'other': '其他' };
    return map[value] || value;
}

function getEntryTypeText(value) {
    const map = { 'single': '一次', 'two': '两次', 'multiple': '多次' };
    return map[value] || value;
}

function getTripPurposeText(purposes) {
    const map = { 'tourism': '旅游', 'business': '商务', 'family': '探亲访友', 
        'culture': '文化', 'sports': '体育', 'study': '学习', 'transit': '过境', 
        'medical': '医疗', 'other': '其他' };
    return purposes ? purposes.map(p => map[p] || p).join(', ') : '-';
}

function getInviterTypeText(value) {
    const map = { 'no': '无（酒店/暂住地）', 'personal': '个人邀请', 'organization': '公司/机构邀请' };
    return map[value] || value;
}

// ============================================
// Word导出 (docx格式)
// ============================================

function exportToWord() {
    const formData = collectFormData();
    const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, BorderStyle } = docx;
    
    // 样式定义
    const headingStyle = {
        font: { name: 'Microsoft YaHei', size: 22, bold: true, color: { hex: '4A6572' } },
        alignment: AlignmentType.CENTER
    };
    
    const sectionTitleStyle = {
        font: { name: 'Microsoft YaHei', size: 13, bold: true, color: { hex: '4A6572' } },
        shading: { fill: 'F0F2F5' }
    };
    
    const labelStyle = {
        font: { name: 'Microsoft YaHei', size: 10, bold: true, color: { hex: '1a1a1a' } },
        shading: { fill: 'F0F2F5' }
    };
    
    const valueStyle = {
        font: { name: 'Microsoft YaHei', size: 10, color: { hex: '1a1a1a' } }
    };
    
    // 创建表格行的辅助函数
    function createTableRow(label, value) {
        return new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: label, ...labelStyle })] })],
                    shading: { fill: 'F0F2F5' },
                    width: { size: 28, type: WidthType.PERCENTAGE }
                }),
                new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: value || '', ...valueStyle })] })],
                    width: { size: 72, type: WidthType.PERCENTAGE }
                })
            ]
        });
    }
    
    // 构建文档 children
    const children = [];
    
    // 标题
    children.push(new Paragraph({
        children: [new TextRun({ text: '申根签证申请表', ...headingStyle })],
        spacing: { after: 100 }
    }));
    
    // 副标题
    children.push(new Paragraph({
        children: [new TextRun({ text: 'Schengen Visa Application Form · 盼达文旅', font: { name: 'Microsoft YaHei', size: 10, color: { hex: '8899A6' } } })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 300 }
    }));
    
    // 一、个人信息
    children.push(new Paragraph({
        children: [new TextRun({ text: '一、个人信息', font: { name: 'Microsoft YaHei', size: 13, bold: true, color: { hex: '4A6572' } } })],
        spacing: { before: 200, after: 100 }
    }));
    
    children.push(new Table({
        rows: [
            createTableRow('姓 (Surname)', formData.surname),
            createTableRow('出生时姓氏', formData.birthSurname),
            createTableRow('名 (Given Name)', formData.givenName),
            createTableRow('出生日期', formatDate(formData.birthDate)),
            createTableRow('出生地', formData.birthPlace),
            createTableRow('出生国家', formData.birthCountry),
            createTableRow('现国籍', formData.nationality),
            createTableRow('性别', formData.gender === 'male' ? '男' : '女'),
            createTableRow('婚姻状况', getMaritalStatusText(formData.maritalStatus)),
            formData.maritalStatus === 'other' ? createTableRow('婚姻状况说明', formData.maritalStatusOther) : null,
            createTableRow('身份证号码', formData.idCard)
        ].filter(Boolean),
        width: { size: 100, type: WidthType.PERCENTAGE }
    }));
    
    // 监护人信息
    if (formData.guardian1Name) {
        children.push(new Paragraph({
            children: [new TextRun({ text: '一.1 监护人信息', font: { name: 'Microsoft YaHei', size: 13, bold: true, color: { hex: '4A6572' } } })],
            spacing: { before: 200, after: 100 }
        }));
        
        children.push(new Table({
            rows:[
                createTableRow('监护人1姓名', formData.guardian1Name),
                createTableRow('监护人1国籍', formData.guardian1Nationality),
                createTableRow('监护人1电话', formData.guardian1Phone),
                createTableRow('监护人1邮箱', formData.guardian1Email),
                createTableRow('监护人1地址', formData.guardian1Address),
                formData.guardian2Name ? createTableRow('监护人2姓名', formData.guardian2Name) : null,
                formData.guardian2Name ? createTableRow('监护人2国籍', formData.guardian2Nationality) : null,
                formData.guardian2Name ? createTableRow('监护人2电话', formData.guardian2Phone) : null,
                formData.guardian2Name ? createTableRow('监护人2邮箱', formData.guardian2Email) : null,
                formData.guardian2Name ? createTableRow('监护人2地址', formData.guardian2Address) : null
            ].filter(Boolean),
            width: { size: 100, type: WidthType.PERCENTAGE }
        }));
    }
    
    // 二、证件与职业信息
    children.push(new Paragraph({
        children: [new TextRun({ text: '二、证件与职业信息', font: { name: 'Microsoft YaHei', size: 13, bold: true, color: { hex: '4A6572' } } })],
        spacing: { before: 200, after: 100 }
    }));
    
    const passportRows = [
        createTableRow('护照种类', getPassportTypeText(formData.passportType)),
        formData.passportType === 'other' ? createTableRow('护照种类说明', formData.passportTypeOther) : null,
        createTableRow('护照号码', formData.passportNumber),
        createTableRow('护照签发日期', formatDate(formData.passportIssueDate)),
        createTableRow('有效期至', formatDate(formData.passportExpiry)),
        createTableRow('签发机关', formData.passportIssuer),
        createTableRow('申请人住址', formData.address),
        createTableRow('电子邮箱', formData.email),
        createTableRow('现职业', getOccupationText(formData.occupation)),
        formData.occupation === 'other' ? createTableRow('职业说明', formData.occupationOther) : null
    ];
    
    if (formData.employerName) {
        passportRows.push(createTableRow('工作单位/学校名称', formData.employerName));
        passportRows.push(createTableRow('工作单位/学校地址', formData.employerAddress));
        passportRows.push(createTableRow('工作单位/学校电话', formData.employerPhone));
    }
    
    children.push(new Table({
        rows: passportRows.filter(Boolean),
        width: { size: 100, type: WidthType.PERCENTAGE }
    }));
    
    // 三、行程信息
    children.push(new Paragraph({
        children: [new TextRun({ text: '三、行程信息', font: { name: 'Microsoft YaHei', size: 13, bold: true, color: { hex: '4A6572' } } })],
        spacing: { before: 200, after: 100 }
    }));
    
    children.push(new Table({
        rows: [
            createTableRow('签证申请国', formData.visaApplicationCountry),
            createTableRow('预计前往申根地区', formData.schengenDestinations),
            createTableRow('申根首入国', formData.firstEntry),
            createTableRow('申请入境次数', getEntryTypeText(formData.entryType)),
            createTableRow('预计入境日期', formatDate(formData.arrivalDate)),
            createTableRow('预计离境日期', formatDate(formData.departureDate)),
            createTableRow('预计逗留天数', formData.stayDuration),
            createTableRow('旅程主要目的', getTripPurposeText(formData.tripPurpose)),
            formData.tripPurpose && formData.tripPurpose.includes('other') ? createTableRow('目的说明', formData.tripPurposeOther) : null
        ].filter(Boolean),
        width: { size: 100, type: WidthType.PERCENTAGE }
    }));
    
    // 四、邀请与住宿信息
    children.push(new Paragraph({
        children: [new TextRun({ text: '四、邀请与住宿信息', font: { name: 'Microsoft YaHei', size: 13, bold: true, color: { hex: '4A6572' } } })],
        spacing: { before: 200, after: 100 }
    }));
    
    let inviterRows = [createTableRow('邀请类型', getInviterTypeText(formData.hasInviter))];
    
    if (formData.hasInviter === 'no') {
        inviterRows.push(createTableRow('酒店/暂住地名称', formData.hotelName));
        inviterRows.push(createTableRow('酒店/暂住地地址', formData.hotelAddress));
        inviterRows.push(createTableRow('酒店/暂住地电话', formData.hotelPhone));
        inviterRows.push(createTableRow('酒店/暂住地邮箱', formData.hotelEmail));
    } else if (formData.hasInviter === 'personal') {
        inviterRows.push(createTableRow('邀请人姓名', formData.inviterName));
        inviterRows.push(createTableRow('邀请人住址', formData.inviterAddress));
        inviterRows.push(createTableRow('邀请人电话', formData.inviterPhone));
        inviterRows.push(createTableRow('邀请人邮箱', formData.inviterEmail));
    } else if (formData.hasInviter === 'organization') {
        inviterRows.push(createTableRow('公司/机构名称', formData.orgName));
        inviterRows.push(createTableRow('公司/机构地址', formData.orgAddress));
        inviterRows.push(createTableRow('公司/机构电话', formData.orgPhone));
        inviterRows.push(createTableRow('联系人姓名', formData.orgContactName));
        inviterRows.push(createTableRow('联系人电话', formData.orgContactPhone));
        inviterRows.push(createTableRow('联系人邮箱', formData.orgContactEmail));
    }
    
    children.push(new Table({
        rows: inviterRows,
        width: { size: 100, type: WidthType.PERCENTAGE }
    }));
    
    // 五、费用与出资信息
    children.push(new Paragraph({
        children: [new TextRun({ text: '五、费用与出资信息', font: { name: 'Microsoft YaHei', size: 13, bold: true, color: { hex: '4A6572' } } })],
        spacing: { before: 200, after: 100 }
    }));
    
    children.push(new Table({
        rows: [
            createTableRow('费用来源', formData.fundingSource === 'applicant' ? '本人支付' : '赞助人支付')
        ],
        width: { size: 100, type: WidthType.PERCENTAGE }
    }));
    
    // 签名区域
    children.push(new Paragraph({
        children: [new TextRun({ text: '申请人签字：_______________________    日期：' + getCurrentDate(), font: { name: 'Microsoft YaHei', size: 10, color: { hex: '1a1a1a' } } })],
        spacing: { before: 400 }
    }));
    
    // 创建文档
    const doc = new Document({
        sections: [{
            properties: {},
            children: children
        }]
    });
    
    // 生成并下载
    Packer.toBlob(doc).then(blob => {
        saveAs(blob, '申根签证申请表.docx');
    });
}
    let content = `
        <html xmlns:o='urn:schemas-microsoft-com:office:office' 
              xmlns:w='urn:schemas-microsoft-com:office:word' 
              xmlns='http://www.w3.org/TR/REC-html40'>
        <head>
            <meta charset='utf-8'>
            <title>申根签证申请表</title>
            <style>
                body { 
                    font-family: 'Microsoft YaHei', '微软雅黑', 'SimSun', Arial, sans-serif; 
                    font-size: 11pt; 
                    line-height: 1.8;
                    color: #1a1a1a;
                    background: #fff;
                    padding: 1.5cm;
                }
                h1 { 
                    text-align: center; 
                    font-size: 22pt; 
                    font-weight: 600;
                    color: #4A6572;
                    margin-bottom: 10px;
                    letter-spacing: 4px;
                }
                .subtitle {
                    text-align: center;
                    font-size: 10pt;
                    color: #8899A6;
                    margin-bottom: 30px;
                    padding-bottom: 20px;
                    border-bottom: 2px solid #B5C1C3;
                }
                h2 { 
                    font-size: 13pt; 
                    font-weight: 600;
                    color: #4A6572;
                    padding: 10px 15px;
                    margin: 25px 0 15px 0;
                    background: linear-gradient(135deg, #E4E6E9 0%, #F0F2F5 100%);
                    border-left: 4px solid #6C8598;
                    border-radius: 0 8px 8px 0;
                }
                h3 {
                    font-size: 11pt;
                    font-weight: 600;
                    color: #5D6D7E;
                    margin: 20px 0 10px 0;
                    padding-bottom: 5px;
                    border-bottom: 1px solid #E4E6E9;
                }
                table { 
                    width: 100%; 
                    border-collapse: collapse; 
                    margin: 15px 0;
                    table-layout: fixed;
                }
                tr:nth-child(even) {
                    background: #F8F9FA;
                }
                td { 
                    border: 1px solid #D5DBDB; 
                    padding: 10px 12px; 
                    word-wrap: break-word;
                }
                td:first-child { 
                    width: 28%;
                    background: #F0F2F5;
                    color: #1a1a1a;
                    font-weight: 600;
                    border-right: none;
                }
                td:last-child {
                    border-left: none;
                }
                .label { 
                    font-weight: 500; 
                    background: #F0F2F5;
                    color: #4A6572;
                }
                .section { 
                    margin-top: 20px; 
                }
                .info-box {
                    background: linear-gradient(135deg, #E4E6E9 0%, #F0F2F5 100%);
                    border: 1px solid #B5C1C3;
                    border-radius: 8px;
                    padding: 15px 20px;
                    margin: 20px 0;
                }
                .info-box p {
                    margin: 5px 0;
                    color: #5D6D7E;
                }
                .signature {
                    margin-top: 50px;
                    padding: 30px;
                    border-top: 2px dashed #B5C1C3;
                }
                .signature-line {
                    display: flex;
                    justify-content: space-between;
                    margin-top: 30px;
                }
                .signature-item {
                    text-align: center;
                    color: #5D6D7E;
                    font-size: 10pt;
                }
                .signature-item span {
                    display: block;
                    margin-top: 8px;
                    color: #8899A6;
                }
                .highlight {
                    color: #6C8598;
                    font-weight: 600;
                }
            </style>
        </head>
        <body>
            <h1>申根签证申请表</h1>
            <p class="subtitle">Schengen Visa Application Form · 盼达文旅</p>
            
            <h2>一、个人信息</h2>
            <table>
                <tr><td class="label">姓 (Surname)</td><td>${formData.surname || ''}</td></tr>
                <tr><td class="label">出生时姓氏</td><td>${formData.birthSurname || ''}</td></tr>
                <tr><td class="label">名 (Given Name)</td><td>${formData.givenName || ''}</td></tr>
                <tr><td class="label">出生日期</td><td>${formatDate(formData.birthDate)}</td></tr>
                <tr><td class="label">出生地</td><td>${formData.birthPlace || ''}</td></tr>
                <tr><td class="label">出生国家</td><td>${formData.birthCountry || ''}</td></tr>
                <tr><td class="label">现国籍</td><td>${formData.nationality || ''}</td></tr>
                <tr><td class="label">性别</td><td>${formData.gender === 'male' ? '男' : '女'}</td></tr>
                <tr><td class="label">婚姻状况</td><td>${getMaritalStatusText(formData.maritalStatus)}</td></tr>
                ${formData.maritalStatus === 'other' ? `<tr><td class="label">婚姻状况说明</td><td>${formData.maritalStatusOther || ''}</td></tr>` : ''}
                <tr><td class="label">身份证号码</td><td>${formData.idCard || ''}</td></tr>
            </table>
    `;
    
    // 监护人信息
    if (formData.guardian1Name) {
        content += `
            <h2>一.1 监护人信息</h2>
            <table>
                <tr><td class="label">监护人1姓名</td><td>${formData.guardian1Name || ''}</td></tr>
                <tr><td class="label">监护人1国籍</td><td>${formData.guardian1Nationality || ''}</td></tr>
                <tr><td class="label">监护人1电话</td><td>${formData.guardian1Phone || ''}</td></tr>
                <tr><td class="label">监护人1邮箱</td><td>${formData.guardian1Email || ''}</td></tr>
                <tr><td class="label">监护人1地址</td><td>${formData.guardian1Address || ''}</td></tr>
                <tr><td class="label">监护人2姓名</td><td>${formData.guardian2Name || ''}</td></tr>
                <tr><td class="label">监护人2国籍</td><td>${formData.guardian2Nationality || ''}</td></tr>
                <tr><td class="label">监护人2电话</td><td>${formData.guardian2Phone || ''}</td></tr>
                <tr><td class="label">监护人2邮箱</td><td>${formData.guardian2Email || ''}</td></tr>
                <tr><td class="label">监护人2地址</td><td>${formData.guardian2Address || ''}</td></tr>
            </table>
        `;
    }
    
    // 证件信息
    content += `
            <h2>二、证件与职业信息</h2>
            <table>
                <tr><td class="label">护照种类</td><td>${getPassportTypeText(formData.passportType)}</td></tr>
                ${formData.passportType === 'other' ? `<tr><td class="label">护照种类说明</td><td>${formData.passportTypeOther || ''}</td></tr>` : ''}
                <tr><td class="label">护照号码</td><td>${formData.passportNumber || ''}</td></tr>
                <tr><td class="label">护照签发日期</td><td>${formatDate(formData.passportIssueDate)}</td></tr>
                <tr><td class="label">有效期至</td><td>${formatDate(formData.passportExpiry)}</td></tr>
                <tr><td class="label">签发机关</td><td>${formData.passportIssuer || ''}</td></tr>
                <tr><td class="label">申请人住址</td><td>${formData.address || ''}</td></tr>
                <tr><td class="label">电子邮箱</td><td>${formData.email || ''}</td></tr>
                <tr><td class="label">现职业</td><td>${getOccupationText(formData.occupation)}</td></tr>
                ${formData.occupation === 'other' ? `<tr><td class="label">职业说明</td><td>${formData.occupationOther || ''}</td></tr>` : ''}
    `;
    
    if (formData.employerName) {
        content += `
                <tr><td class="label">工作单位/学校名称</td><td>${formData.employerName || ''}</td></tr>
                <tr><td class="label">工作单位/学校地址</td><td>${formData.employerAddress || ''}</td></tr>
                <tr><td class="label">工作单位/学校电话</td><td>${formData.employerPhone || ''}</td></tr>
        `;
    }
    
    content += `</table>`;
    
    // 行程信息
    content += `
            <h2>三、行程信息</h2>
            <table>
                <tr><td class="label">签证申请国</td><td>${formData.visaApplicationCountry || ''}</td></tr>
                <tr><td class="label">预计前往申根地区</td><td>${formData.schengenDestinations || ''}</td></tr>
                <tr><td class="label">申根首入国</td><td>${formData.firstEntry || ''}</td></tr>
                <tr><td class="label">申请入境次数</td><td>${getEntryTypeText(formData.entryType)}</td></tr>
                <tr><td class="label">预计入境日期</td><td>${formatDate(formData.arrivalDate)}</td></tr>
                <tr><td class="label">预计离境日期</td><td>${formatDate(formData.departureDate)}</td></tr>
                <tr><td class="label">预计逗留天数</td><td>${formData.stayDuration || ''}</td></tr>
                <tr><td class="label">旅程主要目的</td><td>${getTripPurposeText(formData.tripPurpose)}</td></tr>
                ${formData.tripPurpose && formData.tripPurpose.includes('other') ? `<tr><td class="label">目的说明</td><td>${formData.tripPurposeOther || ''}</td></tr>` : ''}
            </table>
    `;
    
    // 邀请与住宿
    content += `
            <h2>四、邀请与住宿信息</h2>
            <table>
                <tr><td class="label">邀请类型</td><td>${getInviterTypeText(formData.hasInviter)}</td></tr>
    `;
    
    if (formData.hasInviter === 'no') {
        content += `
                <tr><td class="label">酒店/暂住地名称</td><td>${formData.hotelName || ''}</td></tr>
                <tr><td class="label">酒店/暂住地地址</td><td>${formData.hotelAddress || ''}</td></tr>
                <tr><td class="label">酒店/暂住地电话</td><td>${formData.hotelPhone || ''}</td></tr>
                <tr><td class="label">酒店/暂住地邮箱</td><td>${formData.hotelEmail || ''}</td></tr>
        `;
    } else if (formData.hasInviter === 'personal') {
        content += `
                <tr><td class="label">邀请人姓名</td><td>${formData.inviterName || ''}</td></tr>
                <tr><td class="label">邀请人住址</td><td>${formData.inviterAddress || ''}</td></tr>
                <tr><td class="label">邀请人电话</td><td>${formData.inviterPhone || ''}</td></tr>
                <tr><td class="label">邀请人邮箱</td><td>${formData.inviterEmail || ''}</td></tr>
        `;
    } else if (formData.hasInviter === 'organization') {
        content += `
                <tr><td class="label">公司/机构名称</td><td>${formData.orgName || ''}</td></tr>
                <tr><td class="label">公司/机构地址</td><td>${formData.orgAddress || ''}</td></tr>
                <tr><td class="label">公司/机构电话</td><td>${formData.orgPhone || ''}</td></tr>
                <tr><td class="label">联系人姓名</td><td>${formData.orgContactName || ''}</td></tr>
                <tr><td class="label">联系人电话</td><td>${formData.orgContactPhone || ''}</td></tr>
                <tr><td class="label">联系人邮箱</td><td>${formData.orgContactEmail || ''}</td></tr>
        `;
    }
    
    content += `</table>`;
    
    // 费用信息
    content += `
            <h2>五、费用与出资信息</h2>
            <div class="info-box">
                <p><strong>费用来源：</strong><span class="highlight">${formData.fundingSource === 'applicant' ? '本人支付' : '赞助人支付'}</span></p>
            </div>
            
            <div class="signature">
                <div class="signature-line">
                    <div class="signature-item" style="flex: 1;">
                        申请人签字：<span style="border-bottom: 1px solid #333; display: inline-block; min-width: 200px;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
                    </div>
                    <div class="signature-item" style="flex: 1;">
                        日期：<span style="border-bottom: 1px solid #333; display: inline-block; min-width: 150px;">${getCurrentDate()}</span>
                    </div>
                </div>
            </div>
        </body>
        </html>
    `;
    
    // 创建并下载文件
    const blob = new Blob(['\ufeff' + content], { type: 'application/msword;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = '申根签证申请表.doc';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}
