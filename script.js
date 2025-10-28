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

    const downloadEmptyFormButton = document.getElementById('download-empty-form-button');
    const emptyFormDownloadLinks = document.getElementById('empty-form-download-links');

    const defaultFontSize = 13;
    const minFontSize = 7; 

    // 불러온 데이터 건수 표시를 위한 요소 생성 및 추가
    const countDisplay = document.createElement('span');
    countDisplay.textContent = '불러온 데이터 0건';
    countDisplay.style.marginLeft = '10px';
    uploadButton.parentNode.insertBefore(countDisplay, uploadButton.nextSibling);

    let allData = [];
    let currentIndex = 0;

    downloadExcelButton.addEventListener('click', function(event) {
        event.stopPropagation();
        const currentDisplay = excelDownloadLinks.style.display;
        excelDownloadLinks.style.display = currentDisplay === 'none' ? 'flex' : 'none';
        emptyFormDownloadLinks.style.display = 'none';
    });

    downloadEmptyFormButton.addEventListener('click', function(event) {
        event.stopPropagation();
        const currentDisplay = emptyFormDownloadLinks.style.display;
        emptyFormDownloadLinks.style.display = currentDisplay === 'none' ? 'flex' : 'none';
        excelDownloadLinks.style.display = 'none';
    });

    document.addEventListener('click', function(event) {
        if (excelDownloadLinks.style.display === 'flex' && event.target !== downloadExcelButton) {
            if (!excelDownloadLinks.contains(event.target)) {
                excelDownloadLinks.style.display = 'none';
            }
        }
        if (emptyFormDownloadLinks.style.display === 'flex' && event.target !== downloadEmptyFormButton) {
            if (!emptyFormDownloadLinks.contains(event.target)) {
                emptyFormDownloadLinks.style.display = 'none';
            }
        }
    });

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
            uploadButton.disabled = true;
            countDisplay.textContent = '데이터 확인 중...';
            
            setTimeout(() => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    
                    const targetSheetName = workbook.SheetNames.find(name => name.includes('양식'));
                    if (!targetSheetName) {
                        alert('엑셀 파일에 "양식"이라는 이름이 포함된 시트가 없습니다.');
                        resetUI();
                        return;
                    }
                    
                    const worksheet = workbook.Sheets[targetSheetName];
                    
                    const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });

                    const mainHeaderRowIndex = rawData.findIndex(row => row.includes('증권번호'));
                    if (mainHeaderRowIndex === -1) {
                        alert('유효한 헤더를 찾을 수 없습니다.');
                        resetUI();
                        return;
                    }
                    
                    const mainHeaders = rawData[mainHeaderRowIndex];
                    const subHeaders = rawData[mainHeaderRowIndex + 1] || [];
                    
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
                        members.push({
                            type: '주 모집자',
                            name: row[subHeaders.indexOf('이름')], 
                            employeeId: row[subHeaders.indexOf('사원번호')], 
                            contributionRate: row[subHeaders.indexOf('기여율')] 
                        });

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

    
    savePdfButton.addEventListener('click', async function() {
        const element = document.querySelector('.container');
        
        // PDF 변환 직전에 모든 글꼴 크기를 강제로 조정하여 최종 렌더링을 확정합니다.
        document.querySelectorAll('.resizable-text-input').forEach(input => adjustFontSize(input));
        
        // 캔버스 렌더링 직전에 DOM 변화가 확정될 시간을 충분히 확보합니다.
        await new Promise(resolve => setTimeout(resolve, 500)); 

        // PDF 생성 기본 옵션
        const baseOpt = {
            margin: 0,
            image: { type: 'jpeg', quality: 0.98 },
            html2canvas: { 
                scale: 1.2, 
                letterRendering: true 
            },
            jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
        };
        
        // 1. 데이터가 없는 경우 (현재 화면 저장)
        if (allData.length === 0) {
            alert('불러온 데이터가 없습니다. 현재 화면의 빈 양식만 PDF로 저장합니다. 파일 저장 경로를 선택해 주세요.');
            
            const filename = `공동계약확인서_빈양식.pdf`;
            const opt = { ...baseOpt, filename: filename };
            
            await html2pdf().set(opt).from(element).save();
            return;
        }

        // 2. 데이터가 있는 경우 (전체 데이터를 ZIP으로 묶어 저장)
        
        savePdfButton.disabled = true;
        savePdfButton.textContent = 'PDF 파일 생성 중...';
        
        alert(`총 ${allData.length}건의 PDF 파일을 생성하여 ZIP 파일로 묶어 다운로드합니다. 다운로드 경로를 한 번만 지정해주세요.`);
        
        const zip = new JSZip();
        
        for (let i = 0; i < allData.length; i++) {
            displayData(i);
            
            // ⭐️ 핵심 수정: 데이터가 많고 글꼴 조정이 있을 때 안정성을 높이기 위해 1초 대기
            await new Promise(resolve => setTimeout(resolve, 1000)); 

            const contract = allData[i];
            const insuranceCompany = contract.contractDetails.insuranceCompany || '보험사미상';
            const policyNumber = contract.contractDetails.policyNumber || `증권번호미상_${i + 1}`;
            
            const filename = `${insuranceCompany}_${policyNumber}.pdf`.replace(/[\/\?<>\\:\*\|":]/g, '_');
            
            const pdf = await html2pdf().set({ ...baseOpt }).from(element).outputPdf('arraybuffer');
            
            zip.file(filename, pdf);
            
            savePdfButton.textContent = `PDF 생성 중 (${i + 1}/${allData.length})...`;
        }

        const zipFilename = `공동계약확인서_전체_${new Date().toISOString().slice(0, 10)}.zip`;
        zip.generateAsync({ type: 'blob' })
            .then(function(content) {
                const a = document.createElement('a');
                document.body.appendChild(a);
                a.style = 'display: none';
                const url = window.URL.createObjectURL(content);
                a.href = url;
                a.download = zipFilename;
                a.click();
                window.URL.revokeObjectURL(url);
            })
            .finally(() => {
                savePdfButton.disabled = false;
                savePdfButton.textContent = 'PDF로 저장';
            });
        
        alert('ZIP 파일 생성이 완료되었습니다. 다운로드 경로를 선택해 주세요.');
        
        if (allData.length > 0) {
            displayData(currentIndex);
        }
    });

    
    function displayData(index) {
        const contract = allData[index];
        if (!contract) return;

        document.getElementById('insurance-company').value = contract.contractDetails.insuranceCompany || '';
        document.getElementById('policy-number').value = contract.contractDetails.policyNumber || '';
        document.getElementById('contract-date').value = contract.contractDetails.contractDate || '';
        document.getElementById('contractor').value = contract.contractDetails.contractor || '';
        document.getElementById('insured').value = contract.contractDetails.insured || '';
        document.getElementById('monthly-premium').value = contract.contractDetails.monthlyPremium || '';

        const rows = document.getElementById('member-table-body').getElementsByTagName('tr');
        
        for (let i = 0; i < rows.length; i++) {
            const inputs = rows[i].querySelectorAll('input');
            for (let j = 0; j < inputs.length; j++) {
                inputs[j].value = '';
                if (!inputs[j].classList.contains('resizable-text-input')) {
                    inputs[j].style.fontSize = `${defaultFontSize}px`; 
                }
            }
        }

        contract.members.forEach((member, i) => {
            if (rows[i]) {
                const inputs = rows[i].querySelectorAll('input');
                inputs[0].value = member.contributionRate || '';
                inputs[1].value = member.employeeId || '';
                inputs[2].value = member.name || '';
            }
        });

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
        
        document.querySelectorAll('.resizable-text-input').forEach(input => {
            input.style.fontSize = `${defaultFontSize}px`;
        });

        const rows = document.getElementById('member-table-body').getElementsByTagName('tr');
        for (let i = 0; i < rows.length; i++) {
            const inputs = rows[i].querySelectorAll('input');
            for (let j = 0; j < inputs.length; j++) {
                inputs[j].value = '';
                inputs[j].style.fontSize = `${defaultFontSize}px`;
            }
        }

        displayDateElement.textContent = '0000년 00월 00일';
        branchNameInput.value = '##사업단or본부or지점';
        if (branchNameInput2) {
            branchNameInput2.value = '사업단장or본부장or지점장 ###';
        }
        
        prevButton.disabled = true;
        nextButton.disabled = true;
        
        allData = [];
        currentIndex = 0;
        
        excelFileInput.value = '';
        
        countDisplay.textContent = '불러온 데이터 0건';
    }

    function resetUI() {
        uploadButton.disabled = false;
        countDisplay.textContent = '불러온 데이터 0건';
    }

    contractDateInput.addEventListener('change', function() {
        const dateValue = this.value;
        if (dateValue) {
            const dateParts = dateValue.match(/(\d{4}).*?(\d{1,2}).*?(\d{1,2})/);
            
            if (dateParts && dateParts.length === 4) {
                const year = dateParts[1];
                const month = String(dateParts[2]).padStart(2, '0');
                const day = String(dateParts[3]).padStart(2, '0');
                
                if (parseInt(month) >= 1 && parseInt(month) <= 12 && parseInt(day) >= 1 && parseInt(day) <= 31) {
                     displayDateElement.textContent = `${year}년 ${month}월 ${day}일`;
                } else {
                    displayDateElement.textContent = '0000년 00월 00일';
                }
            } else {
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

    function adjustFontSize(inputElement) {
        if (!inputElement.classList.contains('resizable-text-input')) return;

        const parentCell = inputElement.closest('td');
        if (!parentCell) return;
        
        const availableWidth = parentCell.offsetWidth - 20; 
        
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
        
        if (textWidth > availableWidth) {
            let newSize = (availableWidth / textWidth) * defaultFontSize; 
            
            newSize = Math.max(newSize, minFontSize);
            
            inputElement.style.fontSize = `${newSize}px`;
        } else {
            inputElement.style.fontSize = `${defaultFontSize}px`;
        }
    }
    
    const resizableInputs = document.querySelectorAll('.resizable-text-input');
    resizableInputs.forEach(input => {
        input.addEventListener('input', function() {
            adjustFontSize(this);
        });
    });
});
