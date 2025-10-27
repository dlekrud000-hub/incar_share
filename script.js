document.addEventListener('DOMContentLoaded', function() {
    const contractDateInput = document.getElementById('contract-date');
    const displayDateElement = document.getElementById('display-date');
    const branchNameInput = document.getElementById('branch-name');
    const branchNameInput2 = document.getElementById('branch-name-2');
    
    const excelFileInput = document.getElementById('excel-file-input');
    const uploadButton = document.getElementById('upload-button');
    const resetButton = document.getElementById('reset-button');
    const prevButton = document.getElementById('prev-button');
    const nextButton = document.getElementById('next-button');
    const savePdfButton = document.getElementById('save-pdf-button'); 
    
    const downloadExcelButton = document.getElementById('download-excel-button');
    const excelDownloadLinks = document.getElementById('excel-download-links');

    // ⭐️ [추가] 빈양식 다운로드 버튼 및 링크 컨테이너 변수
    const downloadEmptyFormButton = document.getElementById('download-empty-form-button');
    const emptyFormDownloadLinks = document.getElementById('empty-form-download-links');

    // ⭐️ 추가: 엑셀 양식 다운로드 버튼 이벤트 리스너
    downloadExcelButton.addEventListener('click', function(event) {
        event.stopPropagation(); // 버튼 클릭 시 문서 전체 클릭 이벤트 방지
        const currentDisplay = excelDownloadLinks.style.display;
        excelDownloadLinks.style.display = currentDisplay === 'none' ? 'flex' : 'none';
        // 엑셀 버튼 클릭 시 빈양식 링크 숨김
        emptyFormDownloadLinks.style.display = 'none';
    });

    // ⭐️ [추가] 빈양식 다운로드 버튼 이벤트 리스너
    downloadEmptyFormButton.addEventListener('click', function(event) {
        event.stopPropagation(); // 버튼 클릭 시 문서 전체 클릭 이벤트 방지
        const currentDisplay = emptyFormDownloadLinks.style.display;
        emptyFormDownloadLinks.style.display = currentDisplay === 'none' ? 'flex' : 'none';
        // 빈양식 버튼 클릭 시 엑셀 링크 숨김
        excelDownloadLinks.style.display = 'none';
    });

    
    // ⭐️ [수정] 문서의 다른 곳을 클릭하면 다운로드 링크 숨김
    document.addEventListener('click', function(event) {
        if (excelDownloadLinks.style.display === 'flex' && event.target !== downloadExcelButton) {
            if (!excelDownloadLinks.contains(event.target)) {
                excelDownloadLinks.style.display = 'none';
            }
        }
        // ⭐️ [추가] 빈양식 링크 숨김 로직
        if (emptyFormDownloadLinks.style.display === 'flex' && event.target !== downloadEmptyFormButton) {
            if (!emptyFormDownloadLinks.contains(event.target)) {
                emptyFormDownloadLinks.style.display = 'none';
            }
        }
    });

    // 불러온 데이터 건수 표시를 위한 요소 생성 및 추가
    const countDisplay = document.createElement('span');
    countDisplay.textContent = '불러온 데이터 0건';
    countDisplay.style.marginLeft = '10px';
    uploadButton.parentNode.insertBefore(countDisplay, uploadButton.nextSibling);

    let allData = [];
    let currentIndex = 0;

    uploadButton.addEventListener('click', function() {
        excelFileInput.click();
    });

    resetButton.addEventListener('click', function() {
        resetData();
    });

    excelFileInput.addEventListener('change', function(event) {
        const file = event.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = function(e) {
            // 로딩 상태 표시
            uploadButton.disabled = true;
            countDisplay.textContent = '데이터 확인 중...';
            
            setTimeout(() => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    
                    // "양식"이라는 이름이 포함된 시트 찾기
                    const targetSheetName = workbook.SheetNames.find(name => name.includes('양식'));
                    if (!targetSheetName) {
                        alert('엑셀 파일에 "양식"이라는 이름이 포함된 시트가 없습니다.');
                        resetUI();
                        return;
                    }
                    
                    const worksheet = workbook.Sheets[targetSheetName];
                    
                    // 모든 데이터를 배열로 읽기 (헤더 미포함)
                    const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });

                    // 헤더가 있는 행을 찾기. '증권번호'가 포함된 행을 기준으로 한다.
                    const mainHeaderRowIndex = rawData.findIndex(row => row.includes('증권번호'));
                    if (mainHeaderRowIndex === -1) {
                        alert('유효한 헤더를 찾을 수 없습니다.');
                        resetUI();
                        return;
                    }
                    
                    const mainHeaders = rawData[mainHeaderRowIndex];
                    const subHeaders = rawData[mainHeaderRowIndex + 1] || []; // 바로 다음 행을 하위 헤더로 사용
                    
                    // B열(증권번호)에 값이 있는 행만 필터링
                    const validRows = rawData.slice(mainHeaderRowIndex + 2).filter(row => row[mainHeaders.indexOf('증권번호')]);
                    
                    allData = [];
                    
                    validRows.forEach(row => {
                        const contractDetails = {
                            insuranceCompany: row[mainHeaders.indexOf('보험사')],
                            policyNumber: row[mainHeaders.indexOf('증권번호')],
                            contractDate: row[mainHeaders.indexOf('계약일')],
                            contractor: row[mainHeaders.indexOf('계약자')],
                            insured: row[mainHeaders.indexOf('피보험자')],
                            monthlyPremium: row[mainHeaders.indexOf('월 보험료')]
                        };
                        
                        const members = [];
                        // 주모집자 정보
                        members.push({
                            type: '주 모집자',
                            name: row[subHeaders.indexOf('이름')], 
                            employeeId: row[subHeaders.indexOf('사원번호')], 
                            contributionRate: row[subHeaders.indexOf('기여율')] 
                        });

                        // 공동모집자 정보 (최대 5명)
                        for (let i = 1; i <= 5; i++) {
                            const nameIndex = subHeaders.indexOf(`이름(${i})`);
                            if (nameIndex !== -1 && row[nameIndex]) {
                                members.push({
                                    type: `공동모집자(${i})`,
                                    name: row[subHeaders.indexOf(`이름(${i})`)],
                                    employeeId: row[subHeaders.indexOf(`사원번호(${i})`)],
                                    contributionRate: row[subHeaders.indexOf(`기여율(${i})`)]
                                });
                            }
                        }
                        
                        allData.push({
                            contractDetails: contractDetails,
                            members: members
                        });
                    });

                    if (allData.length === 0) {
                        alert('선택한 시트에 유효한 데이터가 없습니다. 빈 양식이 표시됩니다.');
                        prevButton.disabled = true;
                        nextButton.disabled = true;
                    } else {
                        currentIndex = 0;
                        displayData(currentIndex);
                        
                        prevButton.disabled = true;
                        nextButton.disabled = (allData.length <= 1);
                    }

                    // 불러온 데이터 건수 표시 및 알림
                    countDisplay.textContent = `불러온 데이터 ${allData.length}건`;
                    alert(`엑셀 파일 업로드 완료! 총 ${allData.length}건의 데이터를 불러왔습니다.`);
                    
                } catch (error) {
                    alert('파일을 읽는 중 오류가 발생했습니다. 올바른 양식의 엑셀 파일인지 확인해 주세요.');
                    console.error(error);
                    resetUI();
                } finally {
                    uploadButton.disabled = false;
                }
            }, 100); 
        };
        reader.readAsArrayBuffer(file);
    });

    nextButton.addEventListener('click', function() {
        if (currentIndex < allData.length - 1) {
            currentIndex++;
            displayData(currentIndex);
        }
    });

    prevButton.addEventListener('click', function() {
        if (currentIndex > 0) {
            currentIndex--;
            displayData(currentIndex);
        }
    });

    
    // ⭐️ [수정된 부분] PDF 저장 기능 옵션 변경
    savePdfButton.addEventListener('click', async function() {
        const element = document.querySelector('.container');
        
        // PDF 생성 기본 옵션
        const baseOpt = {
            margin: 0,
            image: { type: 'jpeg', quality: 0.98 },
            html2canvas: { 
                scale: 1, // ⭐️ scale 값을 1로 낮춰 A4 용지에 확실하게 맞도록 조정 (핵심)
                letterRendering: true // 텍스트 렌더링 품질 개선
            },
            jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
        };
        
        // 1. 데이터가 없는 경우 (현재 화면 저장)
        if (allData.length === 0) {
            document.querySelectorAll('.resizable-text-input').forEach(input => adjustFontSize(input));
            await new Promise(resolve => setTimeout(resolve, 500)); 

            alert('불러온 데이터가 없습니다. 현재 화면의 빈 양식만 PDF로 저장합니다. 파일 저장 경로를 선택해 주세요.');
            
            const filename = `공동계약확인서_빈양식.pdf`;
            const opt = { ...baseOpt, filename: filename };
            
            await html2pdf().set(opt).from(element).save();
            return;
        }

        // 2. 데이터가 있는 경우 (전체 데이터를 ZIP으로 묶어 저장)
        
        // 버튼 비활성화 및 로딩 메시지
        savePdfButton.disabled = true;
        savePdfButton.textContent = 'PDF 파일 생성 중...';
        
        alert(`총 ${allData.length}건의 PDF 파일을 생성하여 ZIP 파일로 묶어 다운로드합니다. 다운로드 경로를 한 번만 지정해주세요.`);
        
        const zip = new JSZip();
        
        for (let i = 0; i < allData.length; i++) {
            // PDF로 저장할 데이터를 화면에 표시 (이때 displayData 내부에서 글꼴 크기 조정도 호출됨)
            displayData(i);
            
            // DOM이 업데이트될 시간을 잠시 기다림
            await new Promise(resolve => setTimeout(resolve, 500));

            const contract = allData[i];
            const insuranceCompany = contract.contractDetails.insuranceCompany || '보험사미상';
            const policyNumber = contract.contractDetails.policyNumber || `증권번호미상_${i + 1}`;
            
            // 파일 이름을 '보험사_증권번호' 형식으로 설정
            const filename = `${insuranceCompany}_${policyNumber}.pdf`.replace(/[\/\?<>\\:\*\|":]/g, '_');
            
            const pdf = await html2pdf().set({ ...baseOpt }).from(element).outputPdf('arraybuffer');
            
            // ZIP 파일에 PDF 추가
            zip.file(filename, pdf);
            
            savePdfButton.textContent = `PDF 생성 중 (${i + 1}/${allData.length})...`;
        }

        // ZIP 파일 생성 및 다운로드
        const zipFilename = `공동계약확인서_전체_${new Date().toISOString().slice(0, 10)}.zip`;
        zip.generateAsync({ type: 'blob' })
            .then(function(content) {
                const a = document.createElement('a');
                document.body.appendChild(a);
                a.style = 'display: none';
                const url = window.URL.createObjectURL(content);
                a.href = url;
                a.download = zipFilename; // 사용자에게 다운로드 파일명을 제시
                a.click(); // 다운로드 시작 (브라우저가 경로를 묻게 됨)
                window.URL.revokeObjectURL(url);
            })
            .finally(() => {
                savePdfButton.disabled = false;
                savePdfButton.textContent = 'PDF로 저장';
            });
        
        alert('ZIP 파일 생성이 완료되었습니다. 다운로드 경로를 선택해 주세요.');
        
        // 저장이 끝난 후 원래의 데이터로 돌아옴
        if (allData.length > 0) {
            displayData(currentIndex);
        }
    });
    // ---------------------------------------------------

    
    // =========================================================
    // displayData 함수: 데이터 표시 후 모든 resizable-text-input에 글꼴 크기 조정 호출
    // =========================================================
    function displayData(index) {
        const contract = allData[index];
        if (!contract) return;

        // --- 1. 대상 계약 데이터 매핑 ---
        document.getElementById('insurance-company').value = contract.contractDetails.insuranceCompany || '';
        document.getElementById('policy-number').value = contract.contractDetails.policyNumber || '';
        document.getElementById('contract-date').value = contract.contractDetails.contractDate || '';
        document.getElementById('contractor').value = contract.contractDetails.contractor || '';
        document.getElementById('insured').value = contract.contractDetails.insured || '';
        document.getElementById('monthly-premium').value = contract.contractDetails.monthlyPremium || '';

        // --- 2. 계약 기여율 데이터 매핑 ---
        const rows = document.getElementById('member-table-body').getElementsByTagName('tr');
        
        // 모든 행 초기화 및 글꼴 크기 초기화
        for (let i = 0; i < rows.length; i++) {
            const inputs = rows[i].querySelectorAll('input');
            for (let j = 0; j < inputs.length; j++) {
                inputs[j].value = '';
                // 일반 input 필드는 기본 크기로 초기화 (resizable-text-input은 아래에서 다시 처리)
                if (!inputs[j].classList.contains('resizable-text-input')) {
                    inputs[j].style.fontSize = `${defaultFontSize}px`; 
                }
            }
        }

        // 새 데이터로 채우기
        contract.members.forEach((member, i) => {
            if (rows[i]) {
                const inputs = rows[i].querySelectorAll('input');
                inputs[0].value = member.contributionRate || '';
                inputs[1].value = member.employeeId || '';
                inputs[2].value = member.name || ''; // 성명
            }
        });

        // 모든 resizable-text-input 필드에 대해 글꼴 크기 조정 호출
        document.querySelectorAll('.resizable-text-input').forEach(input => {
             adjustFontSize(input);
        });
        
        contractDateInput.dispatchEvent(new Event('change'));
        
        prevButton.disabled = (currentIndex === 0);
        nextButton.disabled = (allData.length <= 1) || (currentIndex === allData.length - 1);
    }
    
    function resetData() {
        document.getElementById('insurance-company').value = '';
        document.getElementById('policy-number').value = '';
        document.getElementById('contract-date').value = '';
        document.getElementById('contractor').value = '';
        document.getElementById('insured').value = '';
        document.getElementById('monthly-premium').value = '';
        
        // 모든 resizable-text-input 필드 글꼴 크기 초기화
        document.querySelectorAll('.resizable-text-input').forEach(input => {
            input.style.fontSize = `${defaultFontSize}px`;
        });

        const rows = document.getElementById('member-table-body').getElementsByTagName('tr');
        for (let i = 0; i < rows.length; i++) {
            const inputs = rows[i].querySelectorAll('input');
            for (let j = 0; j < inputs.length; j++) {
                inputs[j].value = '';
                // 모든 input 필드 글꼴 크기 초기화
                inputs[j].style.fontSize = `${defaultFontSize}px`;
            }
        }

        displayDateElement.textContent = '20\u00A0\u00A0\u00A0\u00A0년\u00A0\u00A0\u00A0\u00A0월\u00A0\u00A0\u00A0\u00A0일';
        branchNameInput.value = '사업단/본부/지점';
        if (branchNameInput2) {
            branchNameInput2.value = '';
        }
        
        // 다음/이전 버튼 상태 초기화
        prevButton.disabled = true;
        nextButton.disabled = true;
        
        // allData 배열 비우기
        allData = [];
        currentIndex = 0;
        
        // 파일 입력 필드 초기화
        excelFileInput.value = '';
        
        // 불러온 데이터 건수 초기화
        countDisplay.textContent = '불러온 데이터 0건';
    }

    // 파일 업로드 실패 또는 유효 데이터가 없을 때 UI를 초기화하는 함수
    function resetUI() {
        uploadButton.disabled = false;
        countDisplay.textContent = '불러온 데이터 0건';
    }

    contractDateInput.addEventListener('change', function() {
        const dateValue = this.value;
        if (dateValue) {
            // 다양한 날짜 형식 처리
            const dateParts = dateValue.match(/(\d{4}).*?(\d{1,2}).*?(\d{1,2})/);
            
            if (dateParts && dateParts.length === 4) {
                const year = dateParts[1];
                const month = String(dateParts[2]).padStart(2, '0');
                const day = String(dateParts[3]).padStart(2, '0');
                
                // 날짜 유효성 검사 (대략적)
                if (parseInt(month) >= 1 && parseInt(month) <= 12 && parseInt(day) >= 1 && parseInt(day) <= 31) {
                     displayDateElement.textContent = `${year}년 ${month}월 ${day}일`;
                } else {
                    displayDateElement.textContent = '0000년 00월 00일';
                }
            } else {
                 // Date 객체를 사용한 기존 로직으로 대체 시도
                 const dateObj = new Date(dateValue);
                 if (!isNaN(dateObj) && dateObj.getFullYear() > 1900) { 
                     const year = dateObj.getFullYear();
                     const month = String(dateObj.getMonth() + 1).padStart(2, '0'); 
                     const day = dateObj.getDate().toString().padStart(2, '0');
                     displayDateElement.textContent = `${year}년 ${month}월 ${day}일`;
                 } else {
                     displayDateElement.textContent = '0000년 00월 00일';
                 }
            }
        } else {
            displayDateElement.textContent = '0000년 00월 00일';
        }
    });

    // 텍스트 길이에 따라 글꼴 크기를 조정하는 함수
    const defaultFontSize = 13;
    const minFontSize = 7; 

    /**
     * 입력 필드의 텍스트 길이에 따라 글꼴 크기를 조정합니다.
     * @param {HTMLInputElement} inputElement - 글꼴 크기를 조정할 input 요소
     */
    function adjustFontSize(inputElement) {
        // resizable-text-input 클래스가 없는 요소는 무시
        if (!inputElement.classList.contains('resizable-text-input')) return;

        const parentCell = inputElement.closest('td');
        if (!parentCell) return;
        
        // 셀의 내부 너비 계산 (td의 padding 10px * 2 = 20px만 제외)
        const availableWidth = parentCell.offsetWidth - 20; 
        
        // 텍스트 너비 측정을 위한 임시 요소 생성
        const tempSpan = document.createElement('span');
        tempSpan.style.fontSize = `${defaultFontSize}px`;
        tempSpan.style.fontFamily = getComputedStyle(inputElement).fontFamily;
        tempSpan.style.whiteSpace = 'nowrap';
        tempSpan.style.position = 'absolute'; 
        tempSpan.style.visibility = 'hidden';
        tempSpan.textContent = inputElement.value;
        document.body.appendChild(tempSpan);
        
        const textWidth = tempSpan.offsetWidth;
        document.body.removeChild(tempSpan); 
        
        // 텍스트 너비가 사용 가능한 너비를 초과하는지 확인 
        if (textWidth > availableWidth) {
            // 초과하면, 비율에 따라 새로운 글꼴 크기를 계산
            let newSize = (availableWidth / textWidth) * defaultFontSize; 
            
            // 최소 크기 제한
            newSize = Math.max(newSize, minFontSize);
            
            // 새로운 글꼴 크기 적용
            inputElement.style.fontSize = `${newSize}px`;
        } else {
            // 초과하지 않으면, 기본(최대) 글꼴 크기로 되돌림
            inputElement.style.fontSize = `${defaultFontSize}px`;
        }
    }
    
    // 모든 resizable-text-input 필드에서 사용자가 직접 입력할 때 글꼴 크기 조정
    const resizableInputs = document.querySelectorAll('.resizable-text-input');
    resizableInputs.forEach(input => {
        input.addEventListener('input', function() {
            adjustFontSize(this);
        });
    });
});


