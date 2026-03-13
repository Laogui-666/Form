/**
 * 申根填表助手 - 智能填表助手
 * 表单逻辑处理与Word导出功能
 */

// 当前步骤
let currentStep = 1;
const totalSteps = 6;

// 表单数据
const formData = {};

// 申根国家列表
const schengenCountries = [
    '法国', '意大利', '西班牙', '德国', '荷兰', '比利时', '奥地利', '瑞士',
    '葡萄牙', '希腊', '丹麦', '瑞典', '挪威', '芬兰', '冰岛', '卢森堡',
    '捷克', '波兰', '斯洛伐克', '斯洛文尼亚', '匈牙利', '立陶宛', '拉脱维亚',
    '爱沙尼亚', '马耳他', '列支敦士登'
];

// 已选目的地
let selectedDestinations = [];

// 初始化
document.addEventListener('DOMContentLoaded', function() {
    initEventListeners();
    initDateInputs();
    updateProgressBar();
    initConditionalLogic();
    setDefaultCheckboxes();
    initDestinationDropdown();
    initStayDurationCalculation();
});

/**
 * 初始化日期选择器 - 使用原生HTML5日期输入
 */
function initDateInputs() {
    // 为所有日期输入框添加事件监听
    const dateInputs = document.querySelectorAll('input[type="date"]');
    dateInputs.forEach(input => {
        input.addEventListener('change', function() {
            // 变化时重新计算逗留天数
            calculateStayDuration();
        });
    });
    
    // 初始化逗留天数计算
    initStayDurationCalculation();
}

/**
 * 初始化目的地多选下拉菜单
 */
function initDestinationDropdown() {
    // 点击外部关闭下拉菜单
    document.addEventListener('click', function(e) {
        const dropdown = document.getElementById('destinationDropdown');
        const wrapper = document.querySelector('.multi-select-wrapper');
        if (wrapper && !wrapper.contains(e.target)) {
            dropdown.classList.remove('show');
        }
    });
}

/**
 * 切换目的地下拉菜单
 */
function toggleDestinationDropdown() {
    const dropdown = document.getElementById('destinationDropdown');
    dropdown.classList.toggle('show');
}

/**
 * 更新目的地选择
 */
function updateDestination() {
    const checkboxes = document.querySelectorAll('#destinationOptions input[type="checkbox"]:checked');
    selectedDestinations = Array.from(checkboxes).map(cb => cb.value);
    
    // 更新输入框
    const input = document.getElementById('destinationInput');
    input.value = selectedDestinations.length > 0 ? `已选择 ${selectedDestinations.length} 个国家` : '';
    
    // 更新标签
    const tagsContainer = document.getElementById('destinationTags');
    tagsContainer.innerHTML = '';
    
    selectedDestinations.forEach(country => {
        const tag = document.createElement('span');
        tag.className = 'selected-tag';
        tag.innerHTML = `${country} <span class="tag-remove" onclick="removeDestination('${country}')">&times;</span>`;
        tagsContainer.appendChild(tag);
    });
}

/**
 * 移除目的地
 */
function removeDestination(country) {
    const checkbox = document.querySelector(`#destinationOptions input[value="${country}"]`);
    if (checkbox) {
        checkbox.checked = false;
        updateDestination();
    }
}

/**
 * 过滤目的地搜索
 */
function filterDestinations() {
    const search = document.getElementById('destinationSearch').value.toLowerCase();
    const options = document.querySelectorAll('.dropdown-option');
    
    options.forEach(option => {
        const text = option.textContent.toLowerCase();
        option.style.display = text.includes(search) ? 'flex' : 'none';
    });
}

/**
 * 初始化事件监听
 */
function initEventListeners() {
    // 生日变化检测 - 未成年人判定 (同时监听change和input事件)
    const birthDateInput = document.getElementById('birthDate');
    if (birthDateInput) {
        birthDateInput.addEventListener('change', checkMinor);
        birthDateInput.addEventListener('input', checkMinor);
    }
    
    // 婚姻状况选择
    document.getElementById('maritalStatus').addEventListener('change', function() {
        toggleField('maritalStatusOtherGroup', this.value === 'other');
    });
    
    // 护照类型选择
    document.getElementById('passportType').addEventListener('change', function() {
        toggleField('passportTypeOtherGroup', this.value === 'other');
    });
    
    // 居留国外 - 选择是时显示居留证种类和有效期
    document.querySelectorAll('input[name="residenceAbroad"]').forEach(radio => {
        radio.addEventListener('change', function() {
            const show = this.value === 'yes';
            toggleField('residencePermitGroup', show);
            toggleField('residencePermitExpiryGroup', show);
        });
    });
    
    // 职业选择
    document.getElementById('occupation').addEventListener('change', function() {
        const showEmployer = ['student', 'employed', 'self-employed'].includes(this.value);
        const employerSection = document.getElementById('employerSection');
        
        if (showEmployer) {
            employerSection.style.display = 'block';
            const label = document.getElementById('employerLabel');
            const nameLabel = document.getElementById('employerNameLabel');
            const positionGroup = document.getElementById('currentPositionGroup');
            
            if (this.value === 'student') {
                label.textContent = '学校信息';
                nameLabel.innerHTML = '学校名称 <span class="required">*</span>';
                positionGroup.style.display = 'none';
                setEmployerRequired(true, false);
            } else if (this.value === 'employed') {
                label.textContent = '工作单位信息';
                nameLabel.innerHTML = '工作单位名称 <span class="required">*</span>';
                positionGroup.style.display = 'flex';
                setEmployerRequired(true, true);
            } else {
                label.textContent = '工作单位信息';
                nameLabel.innerHTML = '工作单位名称 <span class="required">*</span>';
                positionGroup.style.display = 'none';
                setEmployerRequired(true, false);
            }
        } else {
            employerSection.style.display = 'none';
            setEmployerRequired(false, false);
        }
        
        toggleField('occupationOtherGroup', this.value === 'other');
    });
    
    // 旅程目的选择
    document.querySelectorAll('input[name="tripPurpose"]').forEach(checkbox => {
        checkbox.addEventListener('change', function() {
            const otherChecked = Array.from(document.querySelectorAll('input[name="tripPurpose"]'))
                .some(cb => cb.value === 'other' && cb.checked);
            toggleField('tripPurposeOtherGroup', otherChecked);
        });
    });
    
    // 过往申根签证
    document.querySelectorAll('input[name="prevSchengenVisa"]').forEach(radio => {
        radio.addEventListener('change', function() {
            toggleField('prevVisaInfoGroup', this.value === 'yes');
        });
    });
    
    // 指纹记录
    document.querySelectorAll('input[name="fingerprints"]').forEach(radio => {
        radio.addEventListener('change', function() {
            toggleField('fingerprintsDateGroup', this.value === 'yes');
        });
    });
    
    // 邀请人类型
    document.querySelectorAll('input[name="hasInviter"]').forEach(radio => {
        radio.addEventListener('change', function() {
            updateInviterSection(this.value);
        });
    });
    
    // 费用出资来源
    document.querySelectorAll('input[name="fundingSource"]').forEach(radio => {
        radio.addEventListener('change', function() {
            updateFundingSection(this.value);
        });
    });
    
    // 赞助人类型
    document.querySelectorAll('input[name="sponsorType"]').forEach(radio => {
        radio.addEventListener('change', function() {
            toggleField('otherSponsorNameGroup', this.value === 'other');
        });
    });
    
    // 本人支付方式
    document.querySelectorAll('input[name="applicantMeans"]').forEach(checkbox => {
        checkbox.addEventListener('change', function() {
            const otherChecked = Array.from(document.querySelectorAll('input[name="applicantMeans"]'))
                .some(cb => cb.value === 'other' && cb.checked);
            toggleField('applicantMeansOtherGroup', otherChecked);
        });
    });
    
    // 赞助人支付方式
    document.querySelectorAll('input[name="sponsorMeans"]').forEach(checkbox => {
        checkbox.addEventListener('change', function() {
            const otherChecked = Array.from(document.querySelectorAll('input[name="sponsorMeans"]'))
                .some(cb => cb.value === 'other' && cb.checked);
            toggleField('sponsorMeansOtherGroup', otherChecked);
        });
    });
    
    // 同行人
    document.querySelectorAll('input[name="hasCompanion"]').forEach(radio => {
        radio.addEventListener('change', function() {
            toggleCompanionSection(this.value);
        });
    });
}

/**
 * 初始化条件逻辑
 */
function initConditionalLogic() {
    // 初始检查未成年人
    checkMinor();
    
    // 设置默认的本人支付方式（现金和信用卡）
    setDefaultCheckboxes();
}

/**
 * 设置默认勾选
 */
function setDefaultCheckboxes() {
    // 本人支付时，默认勾选现金和信用卡
    setTimeout(() => {
        const applicantCash = document.getElementById('applicantCash');
        const applicantCreditCard = document.getElementById('applicantCreditCard');
        const sponsorAllExpenses = document.getElementById('sponsorAllExpenses');
        
        if (applicantCash) applicantCash.checked = true;
        if (applicantCreditCard) applicantCreditCard.checked = true;
        if (sponsorAllExpenses) sponsorAllExpenses.checked = true;
    }, 100);
}

/**
 * 计算逗留天数 - 供Flatpickr回调调用
 */
function calculateStayDuration() {
    const arrivalDate = document.getElementById('arrivalDate');
    const departureDate = document.getElementById('departureDate');
    const stayDuration = document.getElementById('stayDuration');
    
    if (arrivalDate && departureDate && stayDuration) {
        if (arrivalDate.value && departureDate.value) {
            const arrival = new Date(arrivalDate.value);
            const departure = new Date(departureDate.value);
            
            if (!isNaN(arrival.getTime()) && !isNaN(departure.getTime())) {
                const diffTime = departure - arrival;
                const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
                
                if (diffDays > 0) {
                    stayDuration.value = diffDays + ' 天';
                } else {
                    stayDuration.value = '';
                }
            }
        } else {
            stayDuration.value = '';
        }
    }
}

/**
 * 同步签证申请国到首入国家
 */
function syncVisaCountry() {
    const visaCountry = document.getElementById('visaApplicationCountry');
    const firstEntry = document.getElementById('firstEntry');
    
    if (visaCountry && firstEntry && visaCountry.value) {
        // 检查首入国家选项中是否有对应国家
        let found = false;
        for (let i = 0; i < firstEntry.options.length; i++) {
            if (firstEntry.options[i].value === visaCountry.value) {
                firstEntry.value = visaCountry.value;
                found = true;
                break;
            }
        }
    }
}

// 同行人相关变量
let companionCount = 0;
const companions = [];

/**
 * 切换同行人信息部分显示
 */
function toggleCompanionSection(value) {
    const companionSection = document.getElementById('companionSection');
    if (value === 'yes') {
        companionSection.style.display = 'block';
        // 如果没有同行人，添加一个空的
        if (companions.length === 0) {
            addCompanion();
        }
    } else {
        companionSection.style.display = 'none';
    }
}

/**
 * 添加同行人
 */
function addCompanion() {
    companionCount++;
    const companionId = 'companion_' + companionCount;
    companions.push(companionId);
    
    const companionList = document.getElementById('companionList');
    const companionHtml = `
        <div class="companion-item" id="${companionId}" style="background: rgba(108,133,152,0.1); border-radius: 8px; padding: 15px; margin-bottom: 10px;">
            <div class="form-grid">
                <div class="form-group">
                    <label for="companionName_${companionCount}">同行人姓名 <span class="required">*</span></label>
                    <input type="text" id="companionName_${companionCount}" name="companionName" placeholder="请输入同行人姓名">
                </div>
                <div class="form-group">
                    <label for="companionPassport_${companionCount}">同行人护照号码 <span class="required">*</span></label>
                    <input type="text" id="companionPassport_${companionCount}" name="companionPassport" placeholder="请输入护照号码">
                </div>
                <div class="form-group">
                    <label for="companionRelation_${companionCount}">与同行人关系 <span class="required">*</span></label>
                    <select id="companionRelation_${companionCount}" name="companionRelation" onchange="toggleCompanionRelationOther('${companionCount}', this.value)">
                        <option value="">请选择</option>
                        <option value="spouse">配偶 (Spouse)</option>
                        <option value="parent">父母 (Parent)</option>
                        <option value="child">子女 (Child)</option>
                        <option value="friend">朋友 (Friend)</option>
                        <option value="colleague">同事 (Colleague)</option>
                        <option value="otherRelative">其他亲属 (Other Relative)</option>
                        <option value="other">其他 (Other)</option>
                    </select>
                </div>
                <div class="form-group" id="companionRelationOther_${companionCount}" style="display: none;">
                    <label for="companionRelationText_${companionCount}">请说明关系</label>
                    <input type="text" id="companionRelationText_${companionCount}" name="companionRelationText" placeholder="请具体说明">
                </div>
                <div class="form-group" style="display: flex; align-items: flex-end;">
                    <button type="button" class="btn btn-danger" onclick="removeCompanion('${companionId}')" style="padding: 8px 12px;">
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                            <path d="M3 6h18M19 6v14a2 2 0 01-2 2H7a2 2 0 01-2-2V6m3 0V4a2 2 0 012-2h4a2 2 0 012 2v2"/>
                        </svg>
                        删除
                    </button>
                </div>
            </div>
        </div>
    `;
    companionList.insertAdjacentHTML('beforeend', companionHtml);
}

/**
 * 切换同行人关系其他输入框
 */
function toggleCompanionRelationOther(companionId, value) {
    const otherGroup = document.getElementById('companionRelationOther_' + companionId);
    if (otherGroup) {
        otherGroup.style.display = value === 'other' ? 'flex' : 'none';
    }
}

/**
 * 删除同行人
 */
function removeCompanion(companionId) {
    const companionItem = document.getElementById(companionId);
    if (companionItem) {
        companionItem.remove();
        const index = companions.indexOf(companionId);
        if (index > -1) {
            companions.splice(index, 1);
        }
    }
}

/**
 * 初始化逗留天数自动计算（供兼容性使用）
 */
function initStayDurationCalculation() {
    // 逗留天数计算现在通过原生日期输入的事件监听处理
    // 此函数保留以保持兼容性
}

/**
 * 检查未成年人
 */
function checkMinor() {
    const birthDateInput = document.getElementById('birthDate');
    const guardianSection = document.getElementById('guardianSection');
    
    if (!birthDateInput || !guardianSection) {
        return;
    }
    
    const birthDateValue = birthDateInput.value;
    
    if (birthDateValue) {
        try {
            const birthDate = new Date(birthDateValue);
            const today = new Date();
            
            if (!isNaN(birthDate.getTime())) {
                let age = today.getFullYear() - birthDate.getFullYear();
                const monthDiff = today.getMonth() - birthDate.getMonth();
                
                if (monthDiff < 0 || (monthDiff === 0 && today.getDate() < birthDate.getDate())) {
                    age--;
                }
                
                if (age < 18) {
                    guardianSection.style.display = 'block';
                    document.getElementById('guardian1Name').setAttribute('required', 'required');
                    return;
                }
            }
        } catch (e) {
            console.error('Error parsing birth date:', e);
        }
    }
    
    guardianSection.style.display = 'none';
    document.getElementById('guardian1Name').removeAttribute('required');
}

/**
 * 切换字段显示
 */
function toggleField(fieldId, show) {
    const field = document.getElementById(fieldId);
    if (field) {
        // 修复：隐藏时必须设置 'none' 才能覆盖内联样式
        field.style.display = show ? 'flex' : 'none';
    }
}

/**
 * 设置工作单位/学校信息必填属性
 * @param {boolean} required - 是否必填
 * @param {boolean} showPosition - 是否显示职位字段
 */
function setEmployerRequired(required, showPosition) {
    const employerName = document.getElementById('employerName');
    const employerAddress = document.getElementById('employerAddress');
    const employerPhone = document.getElementById('employerPhone');
    
    if (required) {
        employerName.setAttribute('required', 'required');
        employerAddress.setAttribute('required', 'required');
        employerPhone.setAttribute('required', 'required');
    } else {
        employerName.removeAttribute('required');
        employerAddress.removeAttribute('required');
        employerPhone.removeAttribute('required');
    }
}

/**
 * 更新邀请人部分
 */
function updateInviterSection(value) {
    const hotelSection = document.getElementById('hotelSection');
    const personalSection = document.getElementById('personalInviterSection');
    const orgSection = document.getElementById('orgInviterSection');
    
    hotelSection.style.display = 'none';
    personalSection.style.display = 'none';
    orgSection.style.display = 'none';
    
    if (value === 'no') {
        hotelSection.style.display = 'block';
    } else if (value === 'personal') {
        personalSection.style.display = 'block';
    } else if (value === 'organization') {
        orgSection.style.display = 'block';
    }
}

/**
 * 更新费用出资部分
 */
function updateFundingSection(value) {
    const applicantSection = document.getElementById('applicantFundingSection');
    const sponsorSection = document.getElementById('sponsorFundingSection');
    
    if (value === 'applicant') {
        applicantSection.style.display = 'block';
        sponsorSection.style.display = 'none';
    } else {
        applicantSection.style.display = 'none';
        sponsorSection.style.display = 'block';
    }
}

/**
 * 下一步
 */
function nextStep() {
    console.log('nextStep called, currentStep:', currentStep);
    
    // 验证当前步骤
    if (!validateCurrentStep()) {
        console.log('Validation failed');
        return;
    }
    
    console.log('Validation passed');
    
    // 收集当前步骤数据
    collectCurrentStepData();
    
    if (currentStep < totalSteps) {
        goToStep(currentStep + 1);
    } else {
        // 最后一步，预览
        generatePreview();
    }
}

/**
 * 上一步
 */
function prevStep() {
    if (currentStep > 1) {
        goToStep(currentStep - 1);
    }
}

/**
 * 跳转到指定步骤
 */
function goToStep(step) {
    const currentSection = document.getElementById(`step${currentStep}`);
    const nextSection = document.getElementById(`step${step}`);
    
    // 隐藏当前
    currentSection.classList.remove('active');
    currentSection.style.display = 'none';
    
    // 显示下一步
    nextSection.style.display = 'block';
    nextSection.classList.add('active');
    
    currentStep = step;
    updateProgressBar();
    
    // 如果是预览步骤，生成预览
    if (step === 6) {
        generatePreview();
    }
    
    // 滚动到顶部
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

/**
 * 更新进度条
 */
function updateProgressBar() {
    const progress = (currentStep / totalSteps) * 100;
    document.getElementById('progressFill').style.width = `${progress}%`;
    
    // 更新步骤指示器
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

/**
 * 验证当前步骤
 */
function validateCurrentStep() {
    const currentSection = document.getElementById(`step${currentStep}`);
    if (!currentSection) {
        console.error('Current section not found:', `step${currentStep}`);
        return false;
    }
    
    const requiredInputs = currentSection.querySelectorAll('[required]');
    let isValid = true;
    let firstInvalidInput = null;
    
    requiredInputs.forEach(input => {
        if (!input.value || !input.value.trim()) {
            isValid = false;
            input.style.borderColor = 'var(--error)';
            
            if (!firstInvalidInput) {
                firstInvalidInput = input;
            }
            
            // 恢复边框颜色
            setTimeout(() => {
                input.style.borderColor = '';
            }, 2000);
        }
    });
    
    if (!isValid) {
        showToast('请填写所有必填项', 'error');
        // 滚动到第一个错误输入
        if (firstInvalidInput) {
            firstInvalidInput.scrollIntoView({ behavior: 'smooth', block: 'center' });
            firstInvalidInput.focus();
        }
    }
    
    return isValid;
}

/**
 * 收集当前步骤数据
 */
function collectCurrentStepData() {
    const currentSection = document.getElementById(`step${currentStep}`);
    const inputs = currentSection.querySelectorAll('input, select');
    
    inputs.forEach(input => {
        if (input.type === 'radio' || input.type === 'checkbox') {
            if (input.checked) {
                if (!formData[input.name]) {
                    formData[input.name] = [];
                }
                if (input.type === 'checkbox') {
                    formData[input.name].push(input.value);
                } else {
                    formData[input.name] = input.value;
                }
            }
        } else {
            formData[input.name] = input.value;
        }
    });
    
    // 收集同行人数据
    if (currentStep === 3) {
        const companionData = [];
        const companionItems = document.querySelectorAll('.companion-item');
        companionItems.forEach((item, index) => {
            const nameInput = item.querySelector('[name="companionName"]');
            const passportInput = item.querySelector('[name="companionPassport"]');
            const relationSelect = item.querySelector('[name="companionRelation"]');
            const relationTextInput = item.querySelector('[name="companionRelationText"]');
            
            if (nameInput && nameInput.value) {
                companionData.push({
                    name: nameInput.value,
                    passport: passportInput ? passportInput.value : '',
                    relation: relationSelect ? relationSelect.value : '',
                    relationText: relationTextInput ? relationTextInput.value : ''
                });
            }
        });
        formData.companions = companionData;
    }
}

/**
 * 收集所有数据
 */
function collectAllData() {
    // 清空formData
    for (let key in formData) {
        delete formData[key];
    }
    
    const allSections = document.querySelectorAll('.form-step');
    
    allSections.forEach(section => {
        const inputs = section.querySelectorAll('input, select');
        
        inputs.forEach(input => {
            if (input.type === 'radio' || input.type === 'checkbox') {
                if (input.checked) {
                    if (!formData[input.name]) {
                        formData[input.name] = [];
                    }
                    if (input.type === 'checkbox') {
                        if (!formData[input.name].includes(input.value)) {
                            formData[input.name].push(input.value);
                        }
                    } else {
                        formData[input.name] = input.value;
                    }
                }
            } else {
                formData[input.name] = input.value;
            }
        });
    });
    
    // 处理多选目的地
    formData.destination = selectedDestinations.join(', ');
    
    // 收集同行人数据
    const companionData = [];
    const companionItems = document.querySelectorAll('.companion-item');
    companionItems.forEach((item) => {
        const nameInput = item.querySelector('[name="companionName"]');
        const passportInput = item.querySelector('[name="companionPassport"]');
        const relationSelect = item.querySelector('[name="companionRelation"]');
        const relationTextInput = item.querySelector('[name="companionRelationText"]');
        
        if (nameInput && nameInput.value) {
            companionData.push({
                name: nameInput.value,
                passport: passportInput ? passportInput.value : '',
                relation: relationSelect ? relationSelect.value : '',
                relationText: relationTextInput ? relationTextInput.value : ''
            });
        }
    });
    formData.companions = companionData;
    
    return formData;
}

/**
 * 生成预览
 */
function generatePreview() {
    collectAllData();
    
    const previewContent = document.getElementById('previewContent');
    
    const genderMap = {
        'male': '男 (Male)',
        'female': '女 (Female)'
    };
    
    const maritalMap = {
        'single': '未婚 (Single)',
        'married': '已婚 (Married)',
        'separated': '分居 (Separated)',
        'divorced': '离异 (Divorced)',
        'widowed': '丧偶 (Widowed)',
        'other': '其他 (Other)'
    };
    
    const occupationMap = {
        'student': '学生 (Student)',
        'employed': '在职 (Employed)',
        'self-employed': '自雇 (Self-employed)',
        'retired': '退休 (Retired)',
        'unemployed': '无业 (Unemployed)',
        'other': '其他 (Other)'
    };
    
    const entryTypeMap = {
        'single': '一次 (Single Entry)',
        'two': '两次 (Two Entries)',
        'multiple': '多次 (Multiple Entries)'
    };
    
    const inviterMap = {
        'no': '无 (酒店/暂住地)',
        'personal': '个人邀请',
        'organization': '公司/机构邀请'
    };
    
    const fundingMap = {
        'applicant': '申请人本人支付',
        'sponsor': '赞助人/邀请方支付'
    };
    
    const tripPurposeValues = formData.tripPurpose || [];
    const tripPurposeText = Array.isArray(tripPurposeValues) ? tripPurposeValues.join('、') : tripPurposeValues;
    
    const html = `
        <div class="preview-section">
            <div class="preview-section-title">个人信息</div>
            <div class="preview-item">
                <span class="preview-label">姓</span>
                <span class="preview-value">${formData.surname || '-'}</span>
            </div>
            <div class="preview-item">
                <span class="preview-label">名</span>
                <span class="preview-value">${formData.givenName || '-'}</span>
            </div>
            <div class="preview-item">
                <span class="preview-label">出生日期</span>
                <span class="preview-value">${formData.birthDate || '-'}</span>
            </div>
            <div class="preview-item">
                <span class="preview-label">现国籍</span>
                <span class="preview-value">${formData.nationality || '-'}</span>
            </div>
            <div class="preview-item">
                <span class="preview-label">出生地</span>
                <span class="preview-value">${formData.birthPlace || '-'}</span>
            </div>
            <div class="preview-item">
                <span class="preview-label">出生国家</span>
                <span class="preview-value">${formData.birthCountry || '-'}</span>
            </div>
            <div class="preview-item">
                <span class="preview-label">性别</span>
                <span class="preview-value">${genderMap[formData.gender] || '-'}</span>
            </div>
            <div class="preview-item">
                <span class="preview-label">婚姻状况</span>
                <span class="preview-value">${maritalMap[formData.maritalStatus] || '-'}</span>
            </div>
            ${formData.guardian1Name ? `
            <div class="preview-item">
                <span class="preview-label">监护人1姓名</span>
                <span class="preview-value">${formData.guardian1Name}</span>
            </div>
            ` : ''}
            <div class="preview-item">
                <span class="preview-label">身份证号</span>
                <span class="preview-value">${formData.idCard || '-'}</span>
            </div>
        </div>
        
        <div class="preview-section">
            <div class="preview-section-title">证件与职业信息</div>
            <div class="preview-item">
                <span class="preview-label">护照号码</span>
                <span class="preview-value">${formData.passportNumber || '-'}</span>
            </div>
            <div class="preview-item">
                <span class="preview-label">护照签发日期</span>
                <span class="preview-value">${formData.passportIssueDate || '-'}</span>
            </div>
            <div class="preview-item">
                <span class="preview-label">有效期至</span>
                <span class="preview-value">${formData.passportExpiry || '-'}</span>
            </div>
            <div class="preview-item">
                <span class="preview-label">签发机关</span>
                <span class="preview-value">${formData.passportIssuer || '-'}</span>
            </div>
            <div class="preview-item">
                <span class="preview-label">住址</span>
                <span class="preview-value">${formData.address || '-'}</span>
            </div>
            <div class="preview-item">
                <span class="preview-label">电子邮箱</span>
                <span class="preview-value">${formData.email || '-'}</span>
            </div>
            <div class="preview-item">
                <span class="preview-label">职业</span>
                <span class="preview-value">${occupationMap[formData.occupation] || '-'}</span>
            </div>
            ${formData.currentPosition ? `
            <div class="preview-item">
                <span class="preview-label">当前职位</span>
                <span class="preview-value">${formData.currentPosition}</span>
            </div>
            ` : ''}
        </div>
        
        <div class="preview-section">
            <div class="preview-section-title">行程信息</div>
            <div class="preview-item">
                <span class="preview-label">申根目的地</span>
                <span class="preview-value">${formData.destination || '-'}</span>
            </div>
            <div class="preview-item">
                <span class="preview-label">入境申根国</span>
                <span class="preview-value">${formData.firstEntry || '-'}</span>
            </div>
            <div class="preview-item">
                <span class="preview-label">入境次数</span>
                <span class="preview-value">${entryTypeMap[formData.entryType] || '-'}</span>
            </div>
            <div class="preview-item">
                <span class="preview-label">预计逗留天数</span>
                <span class="preview-value">${formData.stayDuration || '-'}</span>
            </div>
            <div class="preview-item">
                <span class="preview-label">预计入境日期</span>
                <span class="preview-value">${formData.arrivalDate || '-'}</span>
            </div>
            <div class="preview-item">
                <span class="preview-label">预计离境日期</span>
                <span class="preview-value">${formData.departureDate || '-'}</span>
            </div>
            <div class="preview-item">
                <span class="preview-label">旅程目的</span>
                <span class="preview-value">${tripPurposeText || '-'}</span>
            </div>
        </div>
        
        <div class="preview-section">
            <div class="preview-section-title">邀请与住宿</div>
            <div class="preview-item">
                <span class="preview-label">邀请类型</span>
                <span class="preview-value">${inviterMap[formData.hasInviter] || '-'}</span>
            </div>
            ${formData.hotelName ? `
            <div class="preview-item">
                <span class="preview-label">酒店名称</span>
                <span class="preview-value">${formData.hotelName}</span>
            </div>
            ` : ''}
            ${formData.inviterName ? `
            <div class="preview-item">
                <span class="preview-label">邀请人姓名</span>
                <span class="preview-value">${formData.inviterName}</span>
            </div>
            ` : ''}
            ${formData.orgName ? `
            <div class="preview-item">
                <span class="preview-label">机构名称</span>
                <span class="preview-value">${formData.orgName}</span>
            </div>
            ` : ''}
        </div>
        
        <div class="preview-section">
            <div class="preview-section-title">费用与出资</div>
            <div class="preview-item">
                <span class="preview-label">费用来源</span>
                <span class="preview-value">${fundingMap[formData.fundingSource] || '-'}</span>
            </div>
        </div>
    `;
    
    previewContent.innerHTML = html;
}

/**
 * 导出Word文档
 */
function exportToWord() {
    collectAllData();

    // 显示加载状态
    const exportBtn = document.querySelector('.btn-export');
    const originalText = exportBtn.innerHTML;
    exportBtn.classList.add('loading');
    exportBtn.innerHTML = '正在生成...';

    try {
        // 下载HTML格式的Word文档（Word可以打开HTML文件）
        downloadHTMLDocument(formData, '申根签证申请表.doc');
        
        showToast('导出成功！', 'success');
    } catch (error) {
        console.error('导出失败:', error);
        showToast('导出失败，请重试', 'error');
    } finally {
        exportBtn.classList.remove('loading');
        exportBtn.innerHTML = originalText;
    }
}

/**
 * 格式化日期
 */
function formatDate(dateStr) {
    if (!dateStr) return '';
    
    try {
        const date = new Date(dateStr);
        if (isNaN(date.getTime())) return dateStr;
        
        const day = String(date.getDate()).padStart(2, '0');
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const year = date.getFullYear();
        
        return `${day}-${month}-${year}`;
    } catch (e) {
        return dateStr;
    }
}

/**
 * 下载HTML文档（导出为HTML文件，Word可打开）
 */
function downloadHTMLDocument(data, filename) {
    try {
        // 收集所有数据
        collectAllData();
        
        // 创建纯表单格式的内容
        var content = createSimpleFormContent(data);
        
        if (!content || content.length < 100) {
            alert('生成文档内容失败，请重试');
            return;
        }
        
        // 使用简单的a标签下载
        var link = document.createElement('a');
        var blob = new Blob([content], { type: 'text/html;charset=utf-8' });
        link.href = URL.createObjectURL(blob);
        link.download = '申根签证信息表.doc';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        // 清理
        setTimeout(function() {
            URL.revokeObjectURL(link.href);
        }, 100);
        
    } catch (error) {
        alert('导出失败: ' + error.message);
    }
}

/**
 * 创建简单的表单格式内容
 */
function createSimpleFormContent(data) {
    var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>申根签证信息表</title></head><body>';
    html += '<h1 style="text-align:center;color:#2c3e50;">申根签证信息表</h1>';
    html += '<h2 style="color:#4A6572;border-bottom:2px solid #4A6572;">一、个人信息</h2>';
    html += '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse:collapse;width:100%;">';
    
    function addRow(label, value) {
        return '<tr><th style="background:#eee;width:30%;">' + label + '</th><td>' + (value || '-') + '</td></tr>';
    }
    
    var genderText = data.gender === 'male' ? '男' : (data.gender === 'female' ? '女' : '-');
    var maritalText = data.maritalStatus === 'single' ? '未婚' : (data.maritalStatus === 'married' ? '已婚' : (data.maritalStatus === 'other' ? '其他' : '-'));
    
    html += addRow('姓 (Surname)', data.surname);
    html += addRow('名 (Given Name)', data.givenName);
    html += addRow('出生日期', data.birthDate);
    html += addRow('现国籍', data.nationality);
    html += addRow('出生地', data.birthPlace);
    html += addRow('出生国家', data.birthCountry);
    html += addRow('性别', genderText);
    html += addRow('婚姻状况', maritalText);
    html += addRow('身份证号', data.idCard);
    html += '</table>';
    
    if (data.guardian1Name) {
        html += '<h2 style="color:#4A6572;">监护人信息</h2>';
        html += '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse:collapse;width:100%;">';
        html += addRow('监护人1姓名', data.guardian1Name);
        html += addRow('监护人1国籍', data.guardian1Nationality);
        html += addRow('监护人1电话', data.guardian1Phone);
        html += addRow('监护人1邮箱', data.guardian1Email);
        html += addRow('监护人1地址', data.guardian1Address);
        if (data.guardian2Name) {
            html += addRow('监护人2姓名', data.guardian2Name);
            html += addRow('监护人2国籍', data.guardian2Nationality);
            html += addRow('监护人2电话', data.guardian2Phone);
            html += addRow('监护人2邮箱', data.guardian2Email);
            html += addRow('监护人2地址', data.guardian2Address);
        }
        html += '</table>';
    }
    
    html += '<h2 style="color:#4A6572;border-bottom:2px solid #4A6572;">二、证件与职业信息</h2>';
    html += '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse:collapse;width:100%;">';
    
    var passportTypeText = data.passportType === 'ordinary' ? '普通护照' : (data.passportType === 'diplomatic' ? '外交护照' : (data.passportType === 'other' ? '其他' : '-'));
    var occupationText = data.occupation === 'student' ? '学生' : (data.occupation === 'employed' ? '在职' : (data.occupation === 'self-employed' ? '自雇' : (data.occupation === 'retired' ? '退休' : (data.occupation === 'other' ? '其他' : '-'))));
    
    html += addRow('护照号码', data.passportNumber);
    html += addRow('护照种类', passportTypeText);
    html += addRow('签发日期', data.passportIssueDate);
    html += addRow('有效期至', data.passportExpiry);
    html += addRow('签发机关', data.passportIssuer);
    html += addRow('申请人住址', data.address);
    html += addRow('电子邮箱', data.email);
    html += addRow('是否居住在现时国籍以外的国家', data.residenceAbroad === 'yes' ? '是' : '否');
    html += addRow('现职业', occupationText);
    html += addRow('当前职位', data.currentPosition);
    html += '</table>';
    
    if (data.employerName) {
        html += '<h2 style="color:#4A6572;">工作单位/学校信息</h2>';
        html += '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse:collapse;width:100%;">';
        html += addRow('工作单位/学校名称', data.employerName);
        html += addRow('地址', data.employerAddress);
        html += addRow('电话', data.employerPhone);
        html += '</table>';
    }
    
    html += '<h2 style="color:#4A6572;border-bottom:2px solid #4A6572;">三、行程信息</h2>';
    html += '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse:collapse;width:100%;">';
    
    var entryTypeText = data.entryType === 'single' ? '一次' : (data.entryType === 'two' ? '两次' : (data.entryType === 'multiple' ? '多次' : '-'));
    
    html += addRow('签证申请国', data.visaApplicationCountry);
    html += addRow('预计前往申根地区', data.destination);
    html += addRow('申根首入国', data.firstEntry);
    html += addRow('申请入境次数', entryTypeText);
    html += addRow('预计逗留天数', data.stayDuration);
    html += addRow('预计入境日期', data.arrivalDate);
    html += addRow('预计离境日期', data.departureDate);
    html += addRow('旅程主要目的', data.tripPurpose ? data.tripPurpose.join(', ') : '-');
    html += addRow('过去三年是否获批过申根签证', data.prevSchengenVisa === 'yes' ? '是' : '否');
    html += addRow('是否有同行人', data.hasCompanion === 'yes' ? '是' : '否');
    html += '</table>';
    
    html += '<h2 style="color:#4A6572;border-bottom:2px solid #4A6572;">四、邀请与住宿信息</h2>';
    html += '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse:collapse;width:100%;">';
    
    var inviterTypeText = data.hasInviter === 'no' ? '无(酒店/暂住)' : (data.hasInviter === 'personal' ? '个人邀请' : (data.hasInviter === 'organization' ? '机构邀请' : '-'));
    
    html += addRow('邀请类型', inviterTypeText);
    if (data.hotelName) {
        html += addRow('酒店名称', data.hotelName);
        html += addRow('酒店地址', data.hotelAddress);
    }
    if (data.inviterName) {
        html += addRow('邀请人姓名', data.inviterName);
        html += addRow('邀请人地址', data.inviterAddress);
        html += addRow('邀请人电话', data.inviterPhone);
    }
    if (data.orgName) {
        html += addRow('机构名称', data.orgName);
        html += addRow('机构地址', data.orgAddress);
        html += addRow('联系人姓名', data.orgContactName);
        html += addRow('联系人电话', data.orgContactPhone);
    }
    html += '</table>';
    
    html += '<h2 style="color:#4A6572;border-bottom:2px solid #4A6572;">五、费用与出资</h2>';
    html += '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse:collapse;width:100%;">';
    
    var fundingText = data.fundingSource === 'applicant' ? '申请人本人支付' : '赞助人/邀请方支付';
    html += addRow('费用来源', fundingText);
    html += '</table>';
    
    html += '<div style="margin-top:30px;text-align:right;">';
    html += '<p>申请人签名: _______________________</p>';
    html += '<p>日期: ' + new Date().toLocaleDateString('zh-CN') + '</p>';
    html += '</div>';
    
    html += '</body></html>';
    
    return html;
}
    
    // 监护人信息
    if (data.guardian1Name) {
        html += '<table class="section-table">';
        html += '<tr><th colspan="4" class="guardian-title">监护人信息 Guardian Information</th></tr>';
        html += addRow2('监护人1姓名', data.guardian1Name, '监护人1国籍', data.guardian1Nationality);
        html += addRow2('监护人1电话', data.guardian1Phone, '监护人1邮箱', data.guardian1Email);
        html += addRow('监护人1地址', data.guardian1Address);
        if (data.guardian2Name) {
            html += addRow2('监护人2姓名', data.guardian2Name, '监护人2国籍', data.guardian2Nationality);
            html += addRow2('监护人2电话', data.guardian2Phone, '监护人2邮箱', data.guardian2Email);
            html += addRow('监护人2地址', data.guardian2Address);
        }
        html += '</table>';
    }
    
    // 证件与职业信息
    html += '<h2>二、证件与职业信息 Passport & Occupation</h2>';
    html += '<table class="section-table">';
    
    var passportTypeText = data.passportType === 'ordinary' ? '普通护照 (Ordinary Passport)' : 
        (data.passportType === 'diplomatic' ? '外交护照 (Diplomatic Passport)' : 
        (data.passportType === 'service' ? '公务护照 (Service Passport)' : 
        (data.passportType === 'official' ? '因公护照 (Official Passport)' : 
        (data.passportType === 'special' ? '特殊护照 (Special Passport)' : 
        (data.passportType === 'other' ? '其他 (Other)' : '-')))));
    
    var occupationText = data.occupation === 'student' ? '学生 (Student)' : 
        (data.occupation === 'employed' ? '在职 (Employed)' : 
        (data.occupation === 'self-employed' ? '自雇 (Self-employed)' : 
        (data.occupation === 'retired' ? '退休 (Retired)' : 
        (data.occupation === 'unemployed' ? '无业 (Unemployed)' : 
        (data.occupation === 'other' ? '其他 (Other)' : '-')))));
    
    html += addRow2('护照号码 (Passport Number)', data.passportNumber, '护照种类 (Passport Type)', passportTypeText);
    html += addRow2('签发日期 (Issue Date)', data.passportIssueDate, '有效期至 (Expiry Date)', data.passportExpiry);
    html += addRow('签发机关 (Issuing Authority)', data.passportIssuer);
    html += addRow('申请人住址 (Address)', data.address);
    html += addRow('电子邮箱 (Email)', data.email);
    html += addRow('是否居住在现时国籍以外的国家', data.residenceAbroad === 'yes' ? '是 (Yes)' : '否 (No)');
    if (data.residenceAbroad === 'yes') {
        var permitTypeText = data.residencePermitType === 'permanent' ? '永久居留' : 
            (data.residencePermitType === 'temporary' ? '临时居留' : 
            (data.residencePermitType === 'work' ? '工作居留' : 
            (data.residencePermitType === 'student' ? '学习居留' : '其他')));
        html += addRow2('居留证种类', permitTypeText, '居留证有效期', data.residencePermit);
    }
    html += addRow2('现职业 (Occupation)', occupationText, '当前职位', data.currentPosition);
    html += '</table>';
    
    // 工作单位/学校信息
    if (data.employerName) {
        html += '<table class="section-table">';
        html += '<tr><th colspan="4" class="guardian-title">工作单位/学校信息 Employer/School Information</th></tr>';
        html += addRow('工作单位/学校名称', data.employerName);
        html += addRow('地址 (Address)', data.employerAddress);
        html += addRow('电话 (Phone)', data.employerPhone);
        html += '</table>';
    }
    
    // 行程信息
    html += '<h2>三、行程信息 Travel Information</h2>';
    html += '<table class="section-table">';
    
    var entryTypeText = data.entryType === 'single' ? '一次 (Single Entry)' : 
        (data.entryType === 'two' ? '两次 (Two Entries)' : 
        (data.entryType === 'multiple' ? '多次 (Multiple Entries)' : '-'));
    
    html += addRow('签证申请国', data.visaApplicationCountry);
    html += addRow('预计前往申根地区', data.destination);
    html += addRow2('申根首入国 (First Entry Country)', data.firstEntry, '申请入境次数 (Entry Type)', entryTypeText);
    html += addRow2('预计逗留天数 (Duration)', data.stayDuration, '预计入境/离境日期', (data.arrivalDate || '-') + ' 至 ' + (data.departureDate || '-'));
    html += addRow('旅程主要目的 (Purpose)', data.tripPurpose ? data.tripPurpose.join('、') : '-');
    html += addRow('过去三年是否获批过申根签证', data.prevSchengenVisa === 'yes' ? '是 (Yes)' : '否 (No)');
    if (data.prevSchengenVisa === 'yes') {
        html += addRow2('最近一次签证编号', data.prevVisaNumber, '签发日期', data.prevVisaIssueDate);
        html += addRow('有效期至', data.prevVisaExpiryDate);
    }
    html += addRow('以往申请申根签证是否有指纹记录', data.fingerprints === 'yes' ? '是 (Yes)' : '否 (No)');
    if (data.fingerprints === 'yes' && data.fingerprintsDate) {
        html += addRow('指纹记录日期', data.fingerprintsDate);
    }
    html += addRow('是否有同行人', data.hasCompanion === 'yes' ? '是 (Yes)' : '否 (No)');
    html += '</table>';
    
    // 邀请与住宿信息
    html += '<h2>四、邀请与住宿信息 Invitation & Accommodation</h2>';
    html += '<table class="section-table">';
    
    var inviterTypeText = data.hasInviter === 'no' ? '无(酒店/暂住)' : 
        (data.hasInviter === 'personal' ? '个人邀请' : 
        (data.hasInviter === 'organization' ? '机构邀请' : '-'));
    
    html += addRow('邀请类型', inviterTypeText);
    if (data.hotelName) {
        html += addRow2('酒店名称', data.hotelName, '酒店地址', data.hotelAddress);
        html += addRow('酒店电话', data.hotelPhone);
    }
    if (data.inviterName) {
        html += addRow2('邀请人姓名', data.inviterName, '邀请人地址', data.inviterAddress);
        html += addRow('邀请人电话', data.inviterPhone);
    }
    if (data.orgName) {
        html += addRow2('机构名称', data.orgName, '机构地址', data.orgAddress);
        html += addRow2('联系人姓名', data.orgContactName, '联系人电话', data.orgContactPhone);
    }
    html += '</table>';
    
    // 费用与出资信息
    html += '<h2>五、费用与出资 Funding</h2>';
    html += '<table class="section-table">';
    
    var fundingText = data.fundingSource === 'applicant' ? '申请人本人支付' : '赞助人/邀请方支付';
    html += addRow('费用来源', fundingText);
    
    if (data.fundingSource === 'applicant') {
        var means = [];
        if (data.applicantMeans) {
            data.applicantMeans.forEach(function(m) {
                var mText = m === 'cash' ? '现金 (Cash)' : 
                    (m === 'creditCard' ? '信用卡 (Credit Card)' : 
                    (m === 'prepaidAccommodation' ? '预付住宿 (Prepaid Accommodation)' : 
                    (m === 'prepaidTransport' ? '预付交通 (Prepaid Transport)' : 
                    (m === 'other' ? '其他 (Other)' : m)));
                means.push(mText);
            });
        }
        html += addRow('支付方式', means.join('、') || '-');
    } else {
        var sponsorMeans = [];
        if (data.sponsorMeans) {
            data.sponsorMeans.forEach(function(m) {
                var mText = m === 'cash' ? '现金 (Cash)' : 
                    (m === 'accommodationProvided' ? '提供住宿 (Accommodation Provided)' : 
                    (m === 'allExpenses' ? '支付所有费用 (All Expenses)' : 
                    (m === 'prepaidTransport' ? '预付交通 (Prepaid Transport)' : 
                    (m === 'other' ? '其他 (Other)' : m)));
                sponsorMeans.push(mText);
            });
        }
        html += addRow('支付方式', sponsorMeans.join('、') || '-');
        if (data.otherSponsorName) {
            html += addRow('赞助人姓名', data.otherSponsorName);
        }
    }
    html += '</table>';
    
    // 签名区域
    html += '<div class="signature">';
    html += '<div class="signature-line">申请人签名 (Applicant Signature): _______________________</div>';
    html += '<div class="signature-line">日期 (Date): ' + new Date().toLocaleDateString('zh-CN') + '</div>';
    html += '</div>';
    
    html += '<div class="footer-info">';
    html += 'Generated by 盼达文旅 - 申根填表助手';
    html += '</div>';
    
    html += '</body></html>';
    
    return html;
}

/**
 * 创建HTML内容 - 简约高级的表格表单形式
 */
function createHTMLContent(data) {
    const genderMap = { 'male': '男', 'female': '女' };
    const maritalMap = { 'single': '未婚', 'married': '已婚', 'separated': '分居', 'divorced': '离异', 'widowed': '丧偶', 'other': '其他' };
    const occupationMap = { 'student': '学生', 'employed': '在职', 'self-employed': '自雇', 'retired': '退休', 'unemployed': '无业', 'other': '其他' };
    const residencePermitTypeMap = { 'permanent': '永久居留', 'temporary': '临时居留', 'work': '工作居留', 'student': '学习居留', 'other': '其他' };
    const passportTypeMap = { 'ordinary': '普通护照', 'diplomatic': '外交护照', 'service': '公务护照', 'official': '因公护照', 'special': '特殊护照', 'other': '其他' };
    
    const purposes = [];
    document.querySelectorAll('input[name="tripPurpose"]:checked').forEach(cb => {
        const labels = { tourism: '旅游', business: '商务', family: '探亲访友', culture: '文化', sports: '体育', study: '学习', transit: '过境', medical: '医疗', other: '其他' };
        purposes.push(labels[cb.value] || cb.value);
    });
    
    const applicantMeans = [];
    document.querySelectorAll('input[name="applicantMeans"]:checked').forEach(cb => {
        const labels = { cash: '现金', creditCard: '信用卡', prepaidAccommodation: '预付住宿', prepaidTransport: '预付交通', other: '其他' };
        applicantMeans.push(labels[cb.value] || cb.value);
    });
    
    const sponsorMeans = [];
    document.querySelectorAll('input[name="sponsorMeans"]:checked').forEach(cb => {
        const labels = { cash: '现金', accommodationProvided: '提供住宿', allExpenses: '支付所有费用', prepaidTransport: '预付交通', other: '其他' };
        sponsorMeans.push(labels[cb.value] || cb.value);
    });
    
    return `<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="Generator" content="Microsoft Word 15">
<title>申根签证信息表</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: 'Microsoft YaHei', 'SimSun', Arial, sans-serif; font-size: 11pt; line-height: 1.5; color: #1a1a1a; background: #fff; padding: 30px; }
.header { text-align: center; margin-bottom: 30px; padding-bottom: 15px; border-bottom: 3px solid #4A6572; }
.header h1 { font-size: 22pt; color: #2c3e50; margin-bottom: 5px; letter-spacing: 4px; font-weight: bold; }
.header p { font-size: 10pt; color: #5d6d7e; }
.section { margin-bottom: 25px; page-break-inside: avoid; }
.section-title { font-size: 13pt; color: #2c3e50; border-left: 5px solid #4A6572; padding-left: 12px; margin-bottom: 15px; font-weight: bold; background: #e8edef; padding: 10px 12px; 
margin-top: 20px; margin-bottom: 12px; }
table { width: 100%; border-collapse: collapse; margin-bottom: 12px; page-break-inside: avoid; }
table:last-child { margin-bottom: 0; }
th, td { border: 1px solid #5d6d7e; padding: 8px 10px; text-align: left; font-size: 10pt; }
th { background-color: #d5dbe0; color: #1a252f; width: 28%; font-weight: bold; }
td { background-color: #fff; color: #1a1a1a; }
th { font-weight: bold; color: #2c3e50; }
.empty { color: #95a5a6; font-style: italic; }
.footer { margin-top: 40px; padding-top: 25px; border-top: 2px solid #4A6572; display: flex; justify-content: space-between; }
.signature-line { text-align: right; }
.signature-line p { margin: 8px 0; font-size: 10pt; }
</style>
<!--[if gte mso 9]>
<xml>
<w:WordDocument>
<w:View>Print</w:View>
</w:WordDocument>
<![endif]-->
</head>
<body>

<div class="header">
<h1>申根签证信息表</h1>
<p>Schengen Visa Information Form</p>
</div>

<div class="section">
<div class="section-title">一、个人信息 Personal Information</div>
<table>
<tr><th>姓 (Surname)</th><td>${data.surname || '-'}</td><th>名 (Given Name)</th><td>${data.givenName || '-'}</td></tr>
<tr><th>出生日期 (Date of Birth)</th><td>${formatDate(data.birthDate) || '-'}</td><th>现国籍 (Nationality)</th><td>${data.nationality || '-'}</td></tr>
<tr><th>出生地 (Place of Birth)</th><td>${data.birthPlace || '-'}</td><th>出生国家 (Country of Birth)</th><td>${data.birthCountry || '-'}</td></tr>
<tr><th>性别 (Gender)</th><td>${genderMap[data.gender] || '-'}</td><th>婚姻状况 (Marital Status)</th><td>${maritalMap[data.maritalStatus] || '-'}</td></tr>
<tr><th>身份证号 (ID Number)</th><td colspan="3">${data.idCard || '-'}</td></tr>
</table>
${data.guardian1Name ? `
<table>
<tr><th colspan="4" style="background:#ecf0f1;text-align:center;">监护人信息 Guardian Information</th></tr>
<tr><th>监护人1姓名</th><td>${data.guardian1Name || '-'}</td><th>监护人1国籍</th><td>${data.guardian1Nationality || '-'}</td></tr>
<tr><th>监护人1电话</th><td>${data.guardian1Phone || '-'}</td><th>监护人1邮箱</th><td>${data.guardian1Email || '-'}</td></tr>
<tr><th>监护人1地址</th><td colspan="3">${data.guardian1Address || '-'}</td></tr>
${data.guardian2Name ? `
<tr><th>监护人2姓名</th><td>${data.guardian2Name || '-'}</td><th>监护人2国籍</th><td>${data.guardian2Nationality || '-'}</td></tr>
<tr><th>监护人2电话</th><td>${data.guardian2Phone || '-'}</td><th>监护人2邮箱</th><td>${data.guardian2Email || '-'}</td></tr>
<tr><th>监护人2地址</th><td colspan="3">${data.guardian2Address || '-'}</td></tr>
` : ''}
</table>
` : ''}
</div>

<div class="section">
<div class="section-title">二、证件与职业信息 Passport &amp; Occupation</div>
<table>
<tr><th>护照号码 (Passport Number)</th><td>${data.passportNumber || '-'}</td><th>护照种类 (Passport Type)</th><td>${passportTypeMap[data.passportType] || '-'}</td></tr>
<tr><th>签发日期 (Issue Date)</th><td>${formatDate(data.passportIssueDate) || '-'}</td><th>有效期至 (Expiry Date)</th><td>${formatDate(data.passportExpiry) || '-'}</td></tr>
<tr><th>签发机关 (Issuing Authority)</th><td colspan="3">${data.passportIssuer || '-'}</td></tr>
<tr><th>申请人住址 (Address)</th><td colspan="3">${data.address || '-'}</td></tr>
<tr><th>电子邮箱 (Email)</th><td colspan="3">${data.email || '-'}</td></tr>
<tr><th>是否居住在现时国籍以外的国家</th><td colspan="3">${data.residenceAbroad === 'yes' ? '是 (Yes)' : '否 (No)'}</td></tr>
${data.residenceAbroad === 'yes' ? `
<tr><th>居留证种类</th><td>${residencePermitTypeMap[data.residencePermitType] || '-'}</td><th>居留证有效期</th><td>${formatDate(data.residencePermit) || '-'}</td></tr>
` : ''}
<tr><th>现职业 (Occupation)</th><td>${occupationMap[data.occupation] || '-'}</td><th>当前职位</th><td>${data.currentPosition || '-'}</td></tr>
</table>
${data.employerName ? `
<table>
<tr><th colspan="4" style="background:#ecf0f1;text-align:center;">工作单位/学校信息 Employer/School Information</th></tr>
<tr><th>工作单位/学校名称</th><td colspan="3">${data.employerName || '-'}</td></tr>
<tr><th>地址 (Address)</th><td colspan="3">${data.employerAddress || '-'}</td></tr>
<tr><th>电话 (Phone)</th><td colspan="3">${data.employerPhone || '-'}</td></tr>
</table>
` : ''}
</div>

<div class="section">
<div class="section-title">三、行程信息 Travel Information</div>
<table>
<tr><th>签证申请国</th><td colspan="3">${data.visaApplicationCountry || '-'}</td></tr>
<tr><th>预计前往申根地区</th><td colspan="3">${data.destination || '-'}</td></tr>
<tr><th>申根首入国</th><td>${data.firstEntry || '-'}</td><th>申请入境次数 (Entry Type)</th><td>${data.entryType === 'single' ? '一次 (Single)' : data.entryType === 'two' ? '两次 (Two)' : data.entryType === 'multiple' ? '多次 (Multiple)' : '-'}</td></tr>
<tr><th>预计逗留天数 (Duration)</th><td>${data.stayDuration || '-'}</td><th>预计入境/离境日期</th><td>${formatDate(data.arrivalDate) || '-'} 至 ${formatDate(data.departureDate) || '-'}</td></tr>
<tr><th>旅程主要目的 (Purpose)</th><td colspan="3">${purposes.join('、') || '-'}</td></tr>
<tr><th>过去三年是否获批过申根签证</th><td colspan="3">${data.prevSchengenVisa === 'yes' ? '是 (Yes)' : '否 (No)'}</td></tr>
${data.prevSchengenVisa === 'yes' ? `
<tr><th>最近一次签证编号</th><td>${data.prevVisaNumber || '-'}</td><th>签发日期</th><td>${formatDate(data.prevVisaIssueDate) || '-'}</td></tr>
<tr><th>有效期至</th><td colspan="3">${formatDate(data.prevVisaExpiryDate) || '-'}</td></tr>
` : ''}
<tr><th>最近一次申根签证指纹记录</th><td colspan="3">${data.fingerprints === 'yes' ? '是 (Yes)' : '否 (No)'}</td></tr>
${data.fingerprints === 'yes' && data.fingerprintsDate ? `
<tr><th>指纹记录日期</th><td colspan="3">${formatDate(data.fingerprintsDate) || '-'}</td></tr>
` : ''}
<tr><th>是否有同行人</th><td colspan="3">${data.hasCompanion === 'yes' ? '是 (Yes)' : '否 (No)'}</td></tr>
${data.companions && data.companions.length > 0 ? `
<tr><th colspan="4" style="background:#ecf0f1;text-align:center;">同行人信息 Companion Information</th></tr>
<tr><th>序号</th><th>姓名</th><th>护照号码</th><th>关系</th></tr>
${data.companions.map((c, i) => {
    const relationMap = { 'spouse': '配偶', 'parent': '父母', 'child': '子女', 'friend': '朋友', 'colleague': '同事', 'otherRelative': '其他亲属', 'other': c.relationText || '其他' };
    return `<tr><td>${i + 1}</td><td>${c.name || '-'}</td><td>${c.passport || '-'}</td><td>${relationMap[c.relation] || '-'}</td></tr>`;
}).join('')}
` : ''}
</table>
</div>

<div class="section">
<div class="section-title">四、邀请与住宿信息 Invitation &amp; Accommodation</div>
<table>
<tr><th>邀请类型</th><td colspan="3">${data.hasInviter === 'no' ? '无(酒店/暂住)' : data.hasInviter === 'personal' ? '个人邀请' : data.hasInviter === 'organization' ? '机构邀请' : '-'}</td></tr>
${data.hotelName ? `
<tr><th>酒店名称</th><td>${data.hotelName || '-'}</td><th>酒店地址</th><td>${data.hotelAddress || '-'}</td></tr>
` : ''}
${data.inviterName ? `
<tr><th>邀请人姓名</th><td>${data.inviterName || '-'}</td><th>邀请人地址</th><td>${data.inviterAddress || '-'}</td></tr>
<tr><th>邀请人电话</th><td colspan="3">${data.inviterPhone || '-'}</td></tr>
` : ''}
${data.orgName ? `
<tr><th>机构名称</th><td>${data.orgName || '-'}</td><th>机构地址</th><td>${data.orgAddress || '-'}</td></tr>
<tr><th>联系人姓名</th><td>${data.orgContactName || '-'}</td><th>联系人电话</th><td>${data.orgContactPhone || '-'}</td></tr>
` : ''}
</table>
</div>

<div class="section">
<div class="section-title">五、费用与出资 Funding</div>
<table>
<tr><th>费用来源</th><td colspan="3">${data.fundingSource === 'applicant' ? '申请人本人支付' : '赞助人/邀请方支付'}</td></tr>
<tr><th>支付方式</th><td colspan="3">${data.fundingSource === 'applicant' ? applicantMeans.join('、') : sponsorMeans.join('、')}</td></tr>
${data.otherSponsorName ? `
<tr><th>赞助人姓名</th><td colspan="3">${data.otherSponsorName || '-'}</td></tr>
` : ''}
</table>
</div>

<div class="footer">
<div></div>
<div class="signature-line">
<p>申请人签名 (Applicant Signature): _______________________</p>
<p>日期 (Date): ${new Date().toLocaleDateString('zh-CN')}</p>
</div>
</div>

</body>
</html>`;
}

/**
 * 创建RTF内容
 */
function createRTFContent(data) {
    const genderMap = { 'male': '男', 'female': '女' };
    const maritalMap = { 'single': '未婚', 'married': '已婚', 'separated': '分居', 'divorced': '离异', 'widowed': '丧偶', 'other': '其他' };
    const occupationMap = { 'student': '学生', 'employed': '在职', 'self-employed': '自雇', 'retired': '退休', 'unemployed': '无业', 'other': '其他' };
    
    const purposes = [];
    document.querySelectorAll('input[name="tripPurpose"]:checked').forEach(cb => {
        const labels = { tourism: '旅游', business: '商务', family: '探亲访友', culture: '文化', sports: '体育', study: '学习', transit: '过境', medical: '医疗', other: '其他' };
        purposes.push(labels[cb.value] || cb.value);
    });
    
    const applicantMeans = [];
    document.querySelectorAll('input[name="applicantMeans"]:checked').forEach(cb => {
        const labels = { cash: '现金', creditCard: '信用卡', prepaidAccommodation: '预付住宿', prepaidTransport: '预付交通', other: '其他' };
        applicantMeans.push(labels[cb.value] || cb.value);
    });
    
    const sponsorMeans = [];
    document.querySelectorAll('input[name="sponsorMeans"]:checked').forEach(cb => {
        const labels = { cash: '现金', accommodationProvided: '提供住宿', allExpenses: '支付所有费用', prepaidTransport: '预付交通', other: '其他' };
        sponsorMeans.push(labels[cb.value] || cb.value);
    });
    
    const rtf = `{\\rtf1\\ansi\\deff0
{\\fonttbl{\\f0 SimSun;}}
{\\colortbl;\\red0\\green0\\blue0;}
\\paperw12240\\paperh15840
\\margl1800\\margr1800\\margt1440\\margb1440

\\pard\\qc\\fs32\\b 申根签证申请表\\par
\\pard\\qc\\fs20  Schengen Visa Application Form\\par

\\pard\\fs24\\b 一、个人信息\\par
\\pard\\fs18
姓: ${data.surname || ''}\\par
名: ${data.givenName || ''}\\par
出生日期: ${formatDate(data.birthDate)}\\par
现国籍: ${data.nationality || ''}\\par
出生地: ${data.birthPlace || ''}\\par
出生国家: ${data.birthCountry || ''}\\par
性别: ${genderMap[data.gender] || ''}\\par
婚姻状况: ${maritalMap[data.maritalStatus] || ''}\\par
身份证号: ${data.idCard || ''}\\par
${data.guardian1Name ? `监护人1姓名: ${data.guardian1Name}\\par` : ''}
${data.guardian1Phone ? `监护人1电话: ${data.guardian1Phone}\\par` : ''}
${data.guardian1Email ? `监护人1邮箱: ${data.guardian1Email}\\par` : ''}
${data.guardian2Name ? `监护人2姓名: ${data.guardian2Name}\\par` : ''}

\\pard\\fs24\\b 二、证件与职业信息\\par
\\pard\\fs18
护照号码: ${data.passportNumber || ''}\\par
签发日期: ${formatDate(data.passportIssueDate)}\\par
有效期至: ${formatDate(data.passportExpiry)}\\par
签发机关: ${data.passportIssuer || ''}\\par
住址: ${data.address || ''}\\par
电子邮箱: ${data.email || ''}\\par
是否居住在现时国籍以外的国家: ${data.residenceAbroad === 'yes' ? '是' : '否'}\\par
${data.residencePermit ? `居留证有效期: ${formatDate(data.residencePermit)}\\par` : ''}
职业: ${occupationMap[data.occupation] || ''}\\par
${data.currentPosition ? `当前职位: ${data.currentPosition}\\par` : ''}
${data.employerName ? `工作单位/学校: ${data.employerName}\\par` : ''}

\\pard\\fs24\\b三、行程信息\\par
\\pard\\fs18
申根目的地国: ${data.destination || ''}\\par
入境申根国: ${data.firstEntry || ''}\\par
申请入境次数: ${data.entryType || ''}\\par
预计逗留天数: ${data.stayDuration || ''}\\par
预计入境日期: ${formatDate(data.arrivalDate)}\\par
预计离境日期: ${formatDate(data.departureDate)}\\par
旅程目的: ${purposes.join('、')}\\par
过去三年是否获批过申根签证: ${data.prevSchengenVisa === 'yes' ? '是' : '否'}\\par
${data.prevVisaNumber ? `最近一次签证编号: ${data.prevVisaNumber}\\par` : ''}
${data.prevVisaIssueDate ? `最近一次签证签发日期: ${formatDate(data.prevVisaIssueDate)}\\par` : ''}
${data.prevVisaExpiryDate ? `最近一次签证有效期至: ${formatDate(data.prevVisaExpiryDate)}\\par` : ''}
以往申请申根签证是否有指纹记录: ${data.fingerprints === 'yes' ? '是' : '否'}\\par

\\pard\\fs24\\b 四、邀请与住宿信息\\par
\\pard\\fs18
邀请类型: ${data.hasInviter === 'no' ? '无(酒店)' : data.hasInviter === 'personal' ? '个人邀请' : '机构邀请'}\\par
${data.hotelName ? `酒店名称: ${data.hotelName}\\par` : ''}
${data.hotelAddress ? `酒店地址: ${data.hotelAddress}\\par` : ''}
${data.inviterName ? `邀请人姓名: ${data.inviterName}\\par` : ''}
${data.inviterAddress ? `邀请人地址: ${data.inviterAddress}\\par` : ''}
${data.orgName ? `机构名称: ${data.orgName}\\par` : ''}
${data.orgAddress ? `机构地址: ${data.orgAddress}\\par` : ''}
${data.orgContactName ? `联系人姓名: ${data.orgContactName}\\par` : ''}
${data.orgContactPhone ? `联系人电话: ${data.orgContactPhone}\\par` : ''}

\\pard\\fs24\\b 五、费用与出资\\par
\\pard\\fs18
费用来源: ${data.fundingSource === 'applicant' ? '本人支付' : '赞助人支付'}\\par
${data.fundingSource === 'applicant' ? `支付方式: ${applicantMeans.join('、')}` : `支付方式: ${sponsorMeans.join('、')}`}\\par
${data.otherSponsorName ? `赞助人姓名: ${data.otherSponsorName}\\par` : ''}

\\pard\\qc\\fs18
\\par
签名: _______________________  日期: ${new Date().toLocaleDateString('zh-CN')}\\par
}`;
    
    return rtf;
}

/**
 * 显示提示消息
 */
function showToast(message, type = 'info') {
    // 移除已存在的toast
    const existingToast = document.querySelector('.toast');
    if (existingToast) {
        existingToast.remove();
    }
    
    // 创建新的toast
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.textContent = message;
    document.body.appendChild(toast);
    
    // 显示toast
    setTimeout(() => {
        toast.classList.add('show');
    }, 100);
    
    // 隐藏并移除toast
    setTimeout(() => {
        toast.classList.remove('show');
        setTimeout(() => {
            toast.remove();
        }, 300);
    }, 3000);
}
