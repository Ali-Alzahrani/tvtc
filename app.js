
import { db } from './Firebase-config.js';
import { collection, getDocs } from "https://www.gstatic.com/firebasejs/11.2.0/firebase-firestore.js";

//import { uploadData } from './upload.js';

// Attach upload function to the button
// We use this only for one time to upload our data to Firebase
// DO NOT UNCOMMENT THIS UNLESS YOU WANT TO UPLOAD DATA TO FIREBASE
//document.getElementById('uploadButton').addEventListener('click', uploadData);

// Get today's date
const today = new Date();
const formattedDate = today.toLocaleDateString('ar-SA'); // Format it as needed

// Elements
const sectionSelect = document.getElementById('sectionSelect'); // NEW: Section selection
const trainerSelect = document.getElementById('trainerSelect');
const moduleSelect = document.getElementById('moduleSelect');
const downloadButton = document.getElementById('downloadButton');
const day1Select  = document.getElementById('day1Select');
const slot1Select = document.getElementById('slot1Select');
const day2Select  = document.getElementById('day2Select');
const slot2Select = document.getElementById('slot2Select');

const STRATEGY_OPTIONS  = [
    'التدريب بالاكتشاف', 'التدريب البنائي', 'نظريه TRYZ', 
    'حل المشكلات', 'التدريب المعكوس', 'التدريب بالمحاكاه', 
    'دراسه الحاله', 'التدريب المتمايز'
];   // استراتيجيه التدريب (column C)

const EVALUATION_OPTIONS = [
    'لا يوجد', 'واجب', 'مشروع', 'اختبار قصير', 
    'تقييم نظري', 'تقييم عملي', 'اختيار ١', 'اختيار ٢'
]; // اليه التقييم (column B)

// Variable to store selected section
let selectedSection = null;

// When section changes, load the correct trainers and modules
sectionSelect.addEventListener('change', async () => {
    selectedSection = sectionSelect.value;
    if (!selectedSection) {
        trainerSelect.innerHTML = '<option value="">اختر القسم أولاً</option>';
        moduleSelect.innerHTML = '<option value="">اختر القسم أولاً</option>';
        return;
    }

    await Promise.all([loadTrainers(), loadModules()]);
});

// Fetch Trainers from Firebase (depends on section)
async function loadTrainers() {
    let collectionName;
    if (selectedSection === 'قسم الحاسب الالي') {
        collectionName = 'trainers';
    } else if (selectedSection === 'قسم التقنيه الاداريه والماليه') {
        collectionName = 'trainers-administrative-and-financial';
    } else if (selectedSection === 'قسم الالكترونيات') {
        collectionName = 'trainers-electronics';
    } else if (selectedSection === 'قسم التقنيه الكهربائيه') {
        collectionName = 'trainers-electricity';
    }

    const querySnapshot = await getDocs(collection(db, collectionName));
    trainerSelect.innerHTML = '<option value="">اختر المدرب</option>';
    querySnapshot.forEach((doc) => {
        const trainer = doc.data();
        const option = document.createElement('option');
        option.value = JSON.stringify(trainer);
        option.textContent = trainer.name;
        trainerSelect.appendChild(option);
    });
}

// Fetch Modules from Firebase (depends on section)
async function loadModules() {
    let collectionName;
    if (selectedSection === 'قسم الحاسب الالي') {
        collectionName = 'modules-v3';
    } else if (selectedSection === 'قسم التقنيه الاداريه والماليه') {
        collectionName = 'modules-administrative-and-financial';
    } else if (selectedSection === 'قسم الالكترونيات') {
        collectionName = 'modules-electronics';
    } else if (selectedSection === 'قسم التقنيه الكهربائيه') {
        collectionName = 'modules-electricity';
    }

    const querySnapshot = await getDocs(collection(db, collectionName));
    moduleSelect.innerHTML = '<option value="">اختر المقرر</option>';
    querySnapshot.forEach((doc) => {
        const module = doc.data();
        const option = document.createElement('option');
        option.value = JSON.stringify(module);
        option.textContent = module['module_name'];
        moduleSelect.appendChild(option);
    });
}



// ===== NEW: validate the two day+slot selections =====
function validateConnectionHours() {
    const day1  = day1Select?.value;
    const slot1 = slot1Select?.value;
    const day2  = day2Select?.value;
    const slot2 = slot2Select?.value;

    const validDays  = new Set(['sun','mon','tue','wed','thu']);
    const validSlots = new Set(['١ - ٢','٣ - ٤','٥ - ٦']);

    if (!validDays.has(day1) || !validSlots.has(slot1) ||
        !validDays.has(day2) || !validSlots.has(slot2)) {
        alert('الرجاء اختيار يومين وساعات المكتب لكل يوم (1 - 2 أو 3 - 4 أو 5 - 6).');
        return null;
    }
    if (day1 === day2) {
        alert('الرجاء اختيار يومين مختلفين.');
        return null;
    }
    return { day1, slot1, day2, slot2 };
}


// Generate and Download Styled Excel File
downloadButton.addEventListener('click', async () => {
    const selectedTrainer = JSON.parse(trainerSelect.value);
    const selectedModule = JSON.parse(moduleSelect.value);

    if (!selectedTrainer || !selectedModule) {
        alert('Please select both a trainer and a module.');
        return;
    }

    // ===== NEW: ensure the two day+slot fields are valid before generating =====
    const hours = validateConnectionHours();
    if (!hours) return;


    // ######################## The professional template ######################## //
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Module Template');


    // Load the image (from the 'images' folder in your project)
    const response = await fetch('./images/logo.png');
    const imageBlob = await response.blob();  // Renamed from 'blob' to 'imageBlob' to avoid conflict
    const arrayBuffer = await imageBlob.arrayBuffer();
    const imageId = workbook.addImage({
        buffer: arrayBuffer,
        extension: 'png',
    });

    // -------------------  Header Section -----------------------//
    // Add the image
    // Merge 5 cells for the title and 5 cells for the image
    worksheet.mergeCells('A1:E1');  // Merge A1 to E1 for the title
    worksheet.mergeCells('F1:J1');  // Merge F1 to J1 for the image

    worksheet.getRow(1).height = 50;  // Set the height of row 1 to 200



    // Add the image to the merged cells (F1:J1)
    worksheet.addImage(imageId, {
        tl: { col: 5, row: 0 },  // Start at column F (index 5)
        br: { col: 10, row: 1 }  // End at column J (index 10)
    });

    // Add the title in the merged cells (A1:E1)
    const titleCell = worksheet.getCell('A1');
    titleCell.value = "الاداره العامه لجوده التدريب \n ملف المدرب وتوصيف المقرر التدريبي";
    titleCell.font = { size: 16, bold: true };
    titleCell.alignment = { wrapText: true, vertical: 'middle', horizontal: 'center' };

    // Borders - only bottom line -
    // The borders
    worksheet.getCell('A1').border = {
        bottom: { style: 'medium' }
    };
    worksheet.getCell('F1').border = {
        bottom: { style: 'medium' }
    };


    // ----------------  Trainer Information Section ---------------- //
    worksheet.addRow([]);
    worksheet.mergeCells('B3:I3');  // Merge cells from B3 to I3
    const trainerInfoCell = worksheet.getCell('B3');
    trainerInfoCell.value = 'بيانات المدرب';
    trainerInfoCell.font = { bold: true, size: 12 };
    trainerInfoCell.alignment = { vertical: 'middle', horizontal: 'center' };
    trainerInfoCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } };

    // Add text labels
    const cellsToFormat = ['I4', 'I5', 'I6', 'E4', 'E5', 'E6'];
    const labels = ['رقم المدرب', 'اسم المدرب', 'القسم', 'رقم المبنى', 'رقم المكتب', 'الايميل'];

    cellsToFormat.forEach((cellAddress, index) => {
        const cell = worksheet.getCell(cellAddress);
        cell.value = labels[index];
        cell.font = { bold: true, size: 9 };
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
    });

    // Merge cells for input fields
    worksheet.mergeCells('F4:H4');
    worksheet.mergeCells('F5:H5');
    worksheet.mergeCells('F6:H6');
    worksheet.mergeCells('B4:D4');
    worksheet.mergeCells('B5:D5');
    worksheet.mergeCells('B6:D6');

    // Filling the cells with the data coming from Firebase
    worksheet.getCell('F4').value = selectedTrainer.Number || 'N/A';
    worksheet.getCell('F5').value = selectedTrainer.name || 'N/A';
    worksheet.getCell('F6').value = selectedTrainer.section || 'N/A';
    worksheet.getCell('B4').value = selectedTrainer.building_no || 'N/A';
    worksheet.getCell('B5').value = selectedTrainer.office_no || 'N/A';
    worksheet.getCell('B6').value = selectedTrainer.email || 'N/A';


     // Centering the text in all these cells
     const cellsToCenter4 = ["I4", "I5", "I6", "E4", "E5", "E6", "F4", "F5", "F6", "B4", "B5", "B6"]; // List of cells to center
     cellsToCenter4.forEach(cell => {
         worksheet.getCell(cell).alignment = { horizontal: 'center', vertical: 'middle' };
     });

    // Borders
    const borderRange = ['B3', 'I3', 'B4', 'C4', 'D4', 'E4', 'F4', 'G4', 'H4', 'I4',
                        'B5', 'C5', 'D5', 'E5', 'F5', 'G5', 'H5', 'I5',
                        'B6', 'C6', 'D6', 'E6', 'F6', 'G6', 'H6', 'I6'];

    // Apply thin internal borders
    borderRange.forEach(cellAddress => {
        const cell = worksheet.getCell(cellAddress);
        cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
    });

    // Apply medium external borders (manually set for outer edges)
    worksheet.getCell('B3').border.top = { style: 'medium' };
    worksheet.getCell('I3').border.top = { style: 'medium' };

    ['B3', 'B4', 'B5', 'B6'].forEach(cell => worksheet.getCell(cell).border.left = { style: 'medium' });
    ['I3', 'I4', 'I5', 'I6'].forEach(cell => worksheet.getCell(cell).border.right = { style: 'medium' });

    ['B6', 'C6', 'D6', 'E6', 'F6', 'G6', 'H6', 'I6'].forEach(cell => worksheet.getCell(cell).border.bottom = { style: 'medium' });

    // ---------------- Contact Section ---------------- //

    // Trainer Contact  
    worksheet.addRow([]);
    worksheet.mergeCells('B8:I8');
    const contactInfoCell = worksheet.getCell('B8');
    contactInfoCell.value = 'وسيله التواصل';
    contactInfoCell.font = { bold: true, size: 12 };
    contactInfoCell.alignment = { vertical: 'middle', horizontal: 'center' };
    contactInfoCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } };

    worksheet.mergeCells('G9:I13');
    worksheet.getCell('G9').value = 'اليه التواصل مع مدرب المقرر';
    worksheet.getCell('G9').alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getCell('G9').font = { bold: true, size: 12 };

    worksheet.getCell("F9").value = 'الايميل';
    worksheet.getCell("D9").value = 'الساعات المكتبيه';
    worksheet.getCell("B9").value = 'اخرى';

    worksheet.mergeCells('B10:F10');
    worksheet.getCell('B10').value = ('الساعات المكتبيه');

    worksheet.getCell('F11').value = 'الاحد';
    worksheet.getCell('E11').value = 'الاثنين';
    worksheet.getCell('D11').value = 'الثلاثاء';
    worksheet.getCell('C11').value = 'الاربعاء';
    worksheet.getCell('B11').value = 'الخميس';

    worksheet.getCell('F12').value = 'من - الى';
    worksheet.getCell('E12').value = 'من - الى';
    worksheet.getCell('D12').value = 'من - الى';
    worksheet.getCell('C12').value = 'من - الى';
    worksheet.getCell('B12').value = 'من - الى';

    // ===== NEW: Row 13 — office hour slots from the form selections =====
    const dayToCell = {
        sun: 'F13', // الأحد
        mon: 'E13', // الإثنين
        tue: 'D13', // الثلاثاء
        wed: 'C13', // الأربعاء
        thu: 'B13'  // الخميس
    };
    
    // Clear row 13 first
    ['F13','E13','D13','C13','B13'].forEach(addr => { worksheet.getCell(addr).value = ''; });
    
    // Place the two selected slots according to chosen days
    worksheet.getCell(dayToCell[hours.day1]).value = hours.slot1;
    worksheet.getCell(dayToCell[hours.day2]).value = hours.slot2;
    
    // Optional: center these cells like the others
    ['F13','E13','D13','C13','B13'].forEach(addr => {
        worksheet.getCell(addr).alignment = { horizontal: 'center', vertical: 'middle' };
    });
    

    // Centering the text in all these cells
    const cellsToCenter = ["F9", "D9", "B9", "B10", "F11", "E11", "D11", "C11", "B11", "F12", "E12", "D12", "C12", "B12", ]; // List of cells to center
    cellsToCenter.forEach(cell => {
        worksheet.getCell(cell).alignment = { horizontal: 'center', vertical: 'middle' };
    });


    //  Boss Contact 
    worksheet.addRow([]);
    worksheet.addRow([]);
    worksheet.mergeCells('G15:I19');
    worksheet.getCell('G15').value = 'اليه التواصل مع رئيس القسم';
    worksheet.getCell('G15').alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getCell('G15').font = { bold: true, size: 12 };

    worksheet.getCell("F15").value = 'الايميل';
    worksheet.getCell("D15").value = 'الساعات المكتبيه';
    worksheet.getCell("B15").value = 'اخرى';

    worksheet.mergeCells('B16:F16');
    worksheet.getCell('B16').value = ('الساعات المكتبيه');

    worksheet.getCell('F17').value = 'الاحد';
    worksheet.getCell('E17').value = 'الاثنين';
    worksheet.getCell('D17').value = 'الثلاثاء';
    worksheet.getCell('C17').value = 'الاربعاء';
    worksheet.getCell('B17').value = 'الخميس';

    worksheet.getCell('F18').value = 'من - الى';
    worksheet.getCell('E18').value = 'من - الى';
    worksheet.getCell('D18').value = 'من - الى';
    worksheet.getCell('C18').value = 'من - الى';
    worksheet.getCell('B18').value = 'من - الى';

    worksheet.mergeCells('B19:F19');
    worksheet.getCell('B19').value = ('طوال الاسبوع');
    worksheet.getCell('B19').alignment = { horizontal: 'center', vertical: 'middle' };


    // Centering the text in all these cells
    const cellsToCenter2 = ["F15", "D15", "B15", "B16", "F17", "E17", "D17", "C17", "B17", "F18", "E18", "D18", "C18", "B18"]; // List of cells to center
    cellsToCenter2.forEach(cell => {
        worksheet.getCell(cell).alignment = { horizontal: 'center', vertical: 'middle' };
    });


    // Borders
    const contactBorderRanges = [
        ['B8', 'C8', 'D8', 'E8', 'F8', 'G8', 'H8', 'I8',
        'B9', 'C9', 'D9', 'E9', 'F9', 'G9', 'H9', 'I9',
        'B10', 'C10', 'D10', 'E10', 'F10', 'G10', 'H10', 'I10',
        'B11', 'C11', 'D11', 'E11', 'F11', 'G11', 'H11', 'I11',
        'B12', 'C12', 'D12', 'E12', 'F12', 'G12', 'H12', 'I12',
        'B13', 'C13', 'D13', 'E13', 'F13', 'G13', 'H13', 'I13'],
        ['B15', 'C15', 'D15', 'E15', 'F15', 'G15', 'H15', 'I15',
        'B16', 'C16', 'D16', 'E16', 'F16', 'G16', 'H16', 'I16',
        'B17', 'C17', 'D17', 'E17', 'F17', 'G17', 'H17', 'I17',
        'B18', 'C18', 'D18', 'E18', 'F18', 'G18', 'H18', 'I18',
        'B19', 'C19', 'D19', 'E19', 'F19', 'G19', 'H19', 'I19']
    ];

    contactBorderRanges.forEach(borderRange => {
        borderRange.forEach(cellAddress => {
            const cell = worksheet.getCell(cellAddress);
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });
    });

    // Apply medium external borders
    ['B8', 'B9', 'B10', 'B11', 'B12', 'B13'].forEach(cell => worksheet.getCell(cell).border.left = { style: 'medium' });
    ['I8', 'I9', 'I10', 'I11', 'I12', 'I13'].forEach(cell => worksheet.getCell(cell).border.right = { style: 'medium' });
    ['B15', 'B16', 'B17', 'B18', 'B19'].forEach(cell => worksheet.getCell(cell).border.left = { style: 'medium' });
    ['I15', 'I16', 'I17', 'I18', 'I19'].forEach(cell => worksheet.getCell(cell).border.right = { style: 'medium' });
    ['B8', 'C8', 'D8', 'E8', 'F8', 'G8', 'H8', 'I8'].forEach(cell => worksheet.getCell(cell).border.top = { style: 'medium' });
    ['B13', 'C13', 'D13', 'E13', 'F13', 'G13', 'H13', 'I13'].forEach(cell => worksheet.getCell(cell).border.bottom = { style: 'medium' });
    ['B15', 'C15', 'D15', 'E15', 'F15', 'G15', 'H15', 'I15'].forEach(cell => worksheet.getCell(cell).border.top = { style: 'medium' });
    ['B19', 'C19', 'D19', 'E19', 'F19', 'G19', 'H19', 'I19'].forEach(cell => worksheet.getCell(cell).border.bottom = { style: 'medium' });

    
    // ------------------------------  Module Information ------------------------ //
    worksheet.addRow([]);
    worksheet.mergeCells('B21:I21');  // Merge cells from B21 to I21
    const trainerInfoCell2 = worksheet.getCell('B21');
    trainerInfoCell2.value = 'بيانات المقرر التدريبي';
    trainerInfoCell2.font = { bold: true, size: 12 };
    trainerInfoCell2.alignment = { vertical: 'middle', horizontal: 'center' };
    trainerInfoCell2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } };

    // Add text labels
    const cellsToFormat2 = ['H22', 'H23', 'H24', 'H25', 'H26', 'D22', 'D23', 'D24', 'D25', 'D26'];
    const labels2 = ['القسم التدريبي', 'التخصص', 'رمز المقرر', 'اسم المقرر', 'نوع التدريب', 'نمط التدريب', 'مستوى المقرر', 'ساعات الاتصال', 'الساعات المعتمده', 'المتطلب السابق'];

    cellsToFormat2.forEach((cellAddress, index) => {
        const cell = worksheet.getCell(cellAddress);
        cell.value = labels2[index];
        cell.font = { bold: true, size: 9 };
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
    });

    // Merge cells for input fields
    worksheet.mergeCells('F22:G22');
    worksheet.mergeCells('F23:G23');
    worksheet.mergeCells('F24:G24');
    worksheet.mergeCells('F25:G25');
    worksheet.mergeCells('F26:G26');
    worksheet.mergeCells('B22:C22');
    worksheet.mergeCells('B23:C23');
    worksheet.mergeCells('B24:C24');
    worksheet.mergeCells('B25:C25');
    worksheet.mergeCells('B26:C26');
    //
    worksheet.mergeCells('H22:I22');
    worksheet.mergeCells('H23:I23');
    worksheet.mergeCells('H24:I24');
    worksheet.mergeCells('H25:I25');
    worksheet.mergeCells('H26:I26');
    worksheet.mergeCells('D22:E22');
    worksheet.mergeCells('D23:E23');
    worksheet.mergeCells('D24:E24');
    worksheet.mergeCells('D25:E25');
    worksheet.mergeCells('D26:E26');

    // Filling the table with the modules data coming from Firebase
    worksheet.getCell('F24').value = selectedModule.code || 'N/A';
    worksheet.getCell('F25').value = selectedModule.module_name || 'N/A'; 
    worksheet.getCell('B23').value = selectedModule.level || 'N/A';
    worksheet.getCell('B24').value = selectedModule.connection_hours || 'N/A';
    worksheet.getCell('B25').value = selectedModule.approved_hours || 'N/A';
    worksheet.getCell('B26').value = selectedModule.requirement || 'N/A';
    worksheet.getCell('F22').value = selectedModule.section || 'N/A';
    worksheet.getCell('F23').value = selectedModule.major || 'N/A';
    worksheet.getCell('F26').value = selectedModule.training_type || 'N/A';
    worksheet.getCell('B22').value = selectedModule.training_mode || 'N/A';






    // Centering the text in all these cells
    const cellsToCenter3 = ["F22", "F23", "F24", "F25", "F26", "B22", "B23", "B24", "B25", "B26", "H22", "H23", "H24", "H25", "H26", "D22", "D23", "D24", "D25", "D26"]; // List of cells to center
    cellsToCenter3.forEach(cell => {
        worksheet.getCell(cell).alignment = { horizontal: 'center', vertical: 'middle' };
    });

    // Define internal thin borders
    const borderRanges = [
        ['B22', 'C22', 'D22', 'E22', 'F22', 'G22', 'H22', 'I22',
        'B23', 'C23', 'D23', 'E23', 'F23', 'G23', 'H23', 'I23',
        'B24', 'C24', 'D24', 'E24', 'F24', 'G24', 'H24', 'I24',
        'B25', 'C25', 'D25', 'E25', 'F25', 'G25', 'H25', 'I25',
        'B26', 'C26', 'D26', 'E26', 'F26', 'G26', 'H26', 'I26']
    ];
    borderRanges.forEach(borderRange => {
        borderRange.forEach(cellAddress => {
            const cell = worksheet.getCell(cellAddress);
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });
    });
    // Apply medium external borders
    ['B22', 'B23', 'B24', 'B25', 'B26'].forEach(cell => worksheet.getCell(cell).border.left = { style: 'medium' });
    ['I22', 'I23', 'I24', 'I25', 'I26'].forEach(cell => worksheet.getCell(cell).border.right = { style: 'medium' });
    ['B26', 'C26', 'D26', 'E26', 'F26', 'G26', 'H26', 'I26'].forEach(cell => worksheet.getCell(cell).border.bottom = { style: 'medium' });

    // Apply medium external borders for the merged content cell (B21:I31)
    worksheet.getCell('B21').border = {
        top: { style: 'medium' },
        left: { style: 'medium' },
        right: { style: 'medium' },
        bottom: { style: 'thin' }
    };

    // ------------------ Module describtion ------------------- //
    worksheet.mergeCells('B28:I28');
    const moduleDecribtion = worksheet.getCell('B28');
    moduleDecribtion.value = 'وصف المقرر';
    moduleDecribtion.font = { bold: true, size: 12 };
    moduleDecribtion.alignment = { vertical: 'middle', horizontal: 'center' };
    moduleDecribtion.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } };

    worksheet.mergeCells('B29:I32');
    // تعبئة وصف المقرر (B29)
    worksheet.getCell('B29').value =
    selectedModule['وصف المقرر'] || '';
    worksheet.getCell('B29').alignment = { wrapText: true, horizontal: 'right', vertical: 'top' };


    // Apply medium external borders for the merged header cell (B28:I28)
    worksheet.getCell('B28').border = {
        top: { style: 'medium' },
        left: { style: 'medium' },
        right: { style: 'medium' },
        bottom: { style: 'thin' }
    };

    // Apply medium external borders for the merged content cell (B29:I32)
    worksheet.getCell('B29').border = {
        top: { style: 'thin' },
        left: { style: 'medium' },
        right: { style: 'medium' },
        bottom: { style: 'medium' }
    };


    // --------------------------  General Goal ---------------------- //
    worksheet.mergeCells('B34:I34');
    const generalGoal = worksheet.getCell('B34');
    generalGoal.value = 'الهدف العام';
    generalGoal.font = { bold: true, size: 12 };
    generalGoal.alignment = { vertical: 'middle', horizontal: 'center' };
    generalGoal.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } };

    worksheet.mergeCells('B35:I38');
    // تعبئة الهدف العام من المقرر (B35)
    worksheet.getCell('B35').value =
    (selectedModule['الهدف العام من المقرر'] ?? selectedModule['الهدف العام'] ?? '');
    worksheet.getCell('B35').alignment = { wrapText: true, horizontal: 'right', vertical: 'top' };


    // Borders
    worksheet.getCell('B34').border = {
        top: { style: 'medium' },
        left: { style: 'medium' },
        right: { style: 'medium' },
        bottom: { style: 'thin' }
    };

    worksheet.getCell('B35').border = {
        top: { style: 'thin' },
        left: { style: 'medium' },
        right: { style: 'medium' },
        bottom: { style: 'medium' }
    };

    // ---------------- Training requitments ---------------- //
    worksheet.addRow([]);
    worksheet.mergeCells('B40:I40');  
    const trainingRequirement = worksheet.getCell('B40');
    trainingRequirement.value = 'متطلبات التدريب';
    trainingRequirement.font = { bold: true, size: 12 };
    trainingRequirement.alignment = { vertical: 'middle', horizontal: 'center' };
    trainingRequirement.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } };

    // ِEqupments
    worksheet.mergeCells('B41:I41');  
    const equipments = worksheet.getCell('B41');
    equipments.value = 'التجهيزات والخامات';
    equipments.font = { bold: true, size: 12 };
    equipments.alignment = { vertical: 'middle', horizontal: 'center' };
    equipments.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } };

    worksheet.mergeCells('B42:I45');

    // The borders
    worksheet.getCell('B40').border = {
        top: { style: 'medium' },
        left: { style: 'medium' },
        right: { style: 'medium' },
        bottom: { style: 'medium' }
    };

    worksheet.getCell('B41').border = {
        top: { style: 'medium' },
        left: { style: 'medium' },
        right: { style: 'medium' },
        bottom: { style: 'thin' }
    };

    worksheet.getCell('B42').border = {
        top: { style: 'thin' },
        left: { style: 'medium' },
        right: { style: 'medium' },
        bottom: { style: 'medium' }
    };

    // Safty conditions
    worksheet.mergeCells('B47:I47');  
    const safty = worksheet.getCell('B47');
    safty.value = 'تعليمات واشتراطات السلامه';
    safty.font = { bold: true, size: 12 };
    safty.alignment = { vertical: 'middle', horizontal: 'center' };
    safty.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } };

    worksheet.mergeCells('B48:I51');
    // تعبئة تعليمات/اشتراطات السلامة (B48)
    worksheet.getCell('B48').value =
    selectedModule['اشتراطات السلامه'] || '';
    worksheet.getCell('B48').alignment = { wrapText: true, horizontal: 'right', vertical: 'top' };


    // Borders
    worksheet.getCell('B47').border = {
        top: { style: 'medium' },
        left: { style: 'medium' },
        right: { style: 'medium' },
        bottom: { style: 'thin' }
    };

    worksheet.getCell('B48').border = {
        top: { style: 'thin' },
        left: { style: 'medium' },
        right: { style: 'medium' },
        bottom: { style: 'medium' }
    };


    // ---------------- Training plan ---------------- //
    worksheet.mergeCells('A54:J54');  
    const trainingPlan = worksheet.getCell('A54');
    trainingPlan.value = 'الخطه التدريبيه';
    trainingPlan.font = { bold: true, size: 12 };
    trainingPlan.alignment = { vertical: 'middle', horizontal: 'center' };
    trainingPlan.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } };

    // Table header
    worksheet.getCell('J55').value = 'الاسبوع التدريبي'
    worksheet.mergeCells('G55:I55');
    const columnOneTitle = worksheet.getCell('H55');
    columnOneTitle.value = 'الوحدات النظريه والعمليه';
    //worksheet.mergeCells('F55:G55');
    const columnTwoTitle = worksheet.getCell('F55');
    columnTwoTitle.value = 'ساعات التدريب';
    worksheet.getCell('D55').value = 'الاهداف التفصيليه';
    worksheet.mergeCells('D55:E55');
    //worksheet.getCell('E55').value = 'ت'
    //worksheet.getCell('D55').value = 'ساعات التدريب'
    worksheet.getCell('C55').value = 'استراتيجيه التدريب'
    worksheet.getCell('B55').value = 'اليه التقييم'
    worksheet.getCell('A55').value = 'درجه التقييم'

    // Apply alignment and background color to all header cells
    const headerCells = ['J55', 'H55', 'F55', 'E55', 'D55', 'C55', 'B55', 'A55'];
    headerCells.forEach(cellAddress => {
        const cell = worksheet.getCell(cellAddress);
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD3D3D3' } }; // Grey background
        cell.font = { bold: true, size: 8 };
    });

    // Create empty rows 
    for (let i = 1; i <= 57; i++) {
        const rowNumber = 55 + i;
        //const cell = worksheet.getCell(`J${rowNumber}`);
        //cell.value = i; // Fill sequence numbers in column J
        //cell.font = { size: 8 };
        //worksheet.mergeCells(`H${rowNumber}:I${rowNumber}`); // Merge H and I in each row
        //worksheet.mergeCells(`F${rowNumber}:G${rowNumber}`); // Merge F and G in each row
        worksheet.mergeCells(`G${rowNumber}:H${rowNumber}`);
        worksheet.mergeCells(`D${rowNumber}:E${rowNumber}`);
    } 

    // Merge every 3 rows together in column J (rows 56–112)
    // Merge every 3 rows together in column J (rows 56–112) + label weeks vertically
    // ===== Column J: merge every 3 rows into one week + vertical Arabic label =====
    // ===== Column J: merge every 3 rows into one week + vertical Arabic label =====
    const weekLabels = [
        'الأول','الثاني','الثالث','الرابع','الخامس','السادس','السابع','الثامن','التاسع','العاشر',
        'الحادي عشر','الثاني عشر','الثالث عشر','الرابع عشر','الخامس عشر','السادس عشر','السابع عشر','الثامن عشر','التاسع عشر'
    ];
    
    let weekIndex = 0;
    
    for (let startRow = 56; startRow <= 112; startRow += 3) {
        const endRow = Math.min(startRow + 2, 112);
        worksheet.mergeCells(`J${startRow}:J${endRow}`);
    
        const cell = worksheet.getCell(`J${startRow}`);
        cell.value = weekLabels[weekIndex] ?? '';
        weekIndex++;
    
        cell.alignment = {
        vertical: 'middle',
        horizontal: 'center',
        textRotation: 90
        };
        cell.font = { size: 8, bold: true };
    }
    
    // ===== Column I: Arabic-correct lesson/week order =====
    // lesson-week → 1-1, 2-1, 3-1, 1-2, 2-2, ...
    for (let row = 56; row <= 112; row++) {
        const weekNumber   = Math.floor((row - 56) / 3) + 1; // 1..19
        const lessonNumber = ((row - 56) % 3) + 1;           // 1..3
    
        const cell = worksheet.getCell(`I${row}`);
        cell.value = `${lessonNumber}-${weekNumber}`; // ✅ RTL-friendly
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.font = { size: 8 };
    }
  
    // Filling the Excel table with the lessons data
    selectedModule.lessons.forEach((lesson, index) => {
        if (index >= 40) return; // Ensure we don't exceed 40 rows

        const rowNumber = 56 + index; // Rows start from 56
        const lessonCell = worksheet.getCell(`G${rowNumber}`);
        lessonCell.value = lesson; // Assign lesson to the merged column (H:I)        
        
    });

    // ===== NEW: add dropdown data validation for each lesson row =====
    // We add validation only for rows that actually have a lesson.
    const lessonCount = Array.isArray(selectedModule.lessons)
    ? Math.min(40, selectedModule.lessons.length)
    : 0;

    for (let i = 0; i < lessonCount; i++) {
    const row = 56 + i; // lesson rows start at 56

    // Column C: استراتيجيه التدريب
    worksheet.getCell(`C${row}`).dataValidation = {
    type: 'list',
    allowBlank: true,
    formulae: [`"${STRATEGY_OPTIONS.join(',')}"`],
    showInputMessage: true,
    promptTitle: 'اختيار من القائمة',
    prompt: 'اختر قيمة لاستراتيجية التدريب',
    showErrorMessage: true,
    errorTitle: 'قيمة غير صالحة',
    error: 'الرجاء اختيار قيمة من القائمة.'
    };

    // Column B: اليه التقييم
    worksheet.getCell(`B${row}`).dataValidation = {
    type: 'list',
    allowBlank: true,
    formulae: [`"${EVALUATION_OPTIONS.join(',')}"`],
    showInputMessage: true,
    promptTitle: 'اختيار من القائمة',
    prompt: 'اختر قيمة لآلية التقييم',
    showErrorMessage: true,
    errorTitle: 'قيمة غير صالحة',
    error: 'الرجاء اختيار قيمة من القائمة.'
    };
    }

    // The total of training hours and marks
   // The total of training hours and marks

    // The total of training hours and marks

    worksheet.mergeCells('G113:J113');
    const totalHours = worksheet.getCell('G113');
    totalHours.value = 'مجموع ساعات التدريب الفصليه';
    totalHours.font = { bold: true, size: 10 };
    totalHours.alignment = { vertical: 'middle', horizontal: 'center' };
    totalHours.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE8F5E9' }
    };
    totalHours.border = {
        top:    { style: 'medium' },
        left:   { style: 'medium' },
        bottom: { style: 'medium' },
        right:  { style: 'medium' }
    };

    worksheet.mergeCells('B113:E113');
    const totalMarks = worksheet.getCell('B113');
    totalMarks.value = 'مجموع درجات التقييمات من ١٠٠ درجه';
    totalMarks.font = { bold: true, size: 10 };
    totalMarks.alignment = { vertical: 'middle', horizontal: 'center' };
    totalMarks.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE8F5E9' }
    };
    totalMarks.border = {
        top:    { style: 'medium' },
        left:   { style: 'medium' },
        bottom: { style: 'medium' },
        right:  { style: 'medium' }
    };

    // ---- Add bold bottom borders to A113 and F113 ----
    worksheet.getCell('A113').border = {
        bottom: { style: 'medium' }
    };

    worksheet.getCell('F113').border = {
        bottom: { style: 'medium' }
    };

    // The total of training hours and grades
    worksheet.getCell('F113').value = {
        formula: 'SUM(F56:F112)'
    };
    worksheet.getCell('F113').font = { bold: true, size: 10 };
    worksheet.getCell('F113').alignment = { vertical: 'middle', horizontal: 'center' };

    worksheet.getCell('A113').value = {
        formula: 'SUM(A56:A112)'
    };
    worksheet.getCell('A113').font = { bold: true, size: 10 };
    worksheet.getCell('A113').alignment = { vertical: 'middle', horizontal: 'center' };
    
    // Grades summary

    // 1) Merge first
    worksheet.mergeCells('H115:I118');
    worksheet.mergeCells('F115:G116');
    worksheet.mergeCells('D115:E116');
    worksheet.mergeCells('B115:C116');

    worksheet.mergeCells('B117:B118');
    worksheet.mergeCells('C117:C118');
    worksheet.mergeCells('D117:D118');
    worksheet.mergeCells('E117:E118');
    worksheet.mergeCells('F117:F118');
    worksheet.mergeCells('G117:G118');

    // 2) Values
    worksheet.getCell('H115').value = 'تنويه: يمكن التعرف على طرق ووسائل التدريب من خلال الرابط';
    worksheet.getCell('F115').value = 'درجه الاعمال الفصليه من ٦٠';
    worksheet.getCell('D115').value = 'درجه الاختبار النهائي من ٤٠';
    worksheet.getCell('B115').value = 'مجموع الدرجات من ١٠٠';

    worksheet.getCell('G117').value = 0;
    worksheet.getCell('E117').value = 0;
    worksheet.getCell('C117').value = 0;

    // Red background for the three cells
    ['G117', 'E117', 'C117'].forEach(addr => {
        worksheet.getCell(addr).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFC7CE' } // soft red (Excel-style warning)
        };
    });


    worksheet.getCell('F117').value = 'من ٦٠';
    worksheet.getCell('D117').value = 'من ٤٠';
    worksheet.getCell('B117').value = 'من ١٠٠';

    // 3) Wrap + Center + Font
    const cellsToCenterWrapBold = ['H115', 'F115', 'D115', 'B115'];
    cellsToCenterWrapBold.forEach(addr => {
    const cell = worksheet.getCell(addr);
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.font = { bold: true, size: 10 };
    });

    const cellsToCenterWrap = ['B117', 'C117', 'D117', 'E117', 'F117', 'G117'];
    cellsToCenterWrap.forEach(addr => {
    const cell = worksheet.getCell(addr);
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.font = { size: 10 };
    });


    // 4) Borders: thin for all cells in B115:I118
    for (let r = 115; r <= 118; r++) {
    ['B','C','D','E','F','G','H','I'].forEach(col => {
        const cell = worksheet.getCell(`${col}${r}`);
        cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
        };
    });
    }

    // 5) Medium outer border around B115:I118
    // Top (row 115)
    ['B','C','D','E','F','G','H','I'].forEach(col => {
    worksheet.getCell(`${col}115`).border.top = { style: 'medium' };
    });

    // Bottom (row 118)
    ['B','C','D','E','F','G','H','I'].forEach(col => {
    worksheet.getCell(`${col}118`).border.bottom = { style: 'medium' };
    });

    // Left edge (col B)
    [115,116,117,118].forEach(r => {
    worksheet.getCell(`B${r}`).border.left = { style: 'medium' };
    });

    // Right edge (col I)
    [115,116,117,118].forEach(r => {
    worksheet.getCell(`I${r}`).border.right = { style: 'medium' };
    });



    // Borders
    const tableRange = [];
    for (let i = 55; i <= 112; i++) {
        ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'].forEach(col => {
            tableRange.push(`${col}${i}`);
        });
    }

    tableRange.forEach(cellAddress => {
        const cell = worksheet.getCell(cellAddress);
        cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
    });

    // Apply medium external borders
    ['A55', 'B55', 'C55', 'D55', 'E55', 'F55', 'G55', 'H55', 'I55', 'J55'].forEach(cell => worksheet.getCell(cell).border.top = { style: 'medium' });
    ['A112', 'B112', 'C112', 'D112', 'E112', 'F112', 'G112', 'H112', 'I112', 'J112'].forEach(cell => worksheet.getCell(cell).border.bottom = { style: 'medium' });
    ['A55', 'A56', 'A57', 'A58', 'A59', 'A60', 'A61', 'A62', 'A63', 'A64', 'A65', 'A66', 'A67', 'A68', 'A69', 'A70', 'A71', 'A72', 'A73', 'A74', 'A75', 'A76', 'A77', 'A78', 'A79', 'A80', 'A81', 'A82', 'A83', 'A84', 'A85', 'A86', 'A87', 'A88', 'A89', 'A90', 'A91', 'A92', 'A93', 'A94', 'A95', 'A96', 'A97', 'A98', 'A99', 'A100','A101', 'A102', 'A103', 'A104', 'A105', 'A106','A107', 'A108', 'A109', 'A110', 'A111', 'A112'].forEach(cell => worksheet.getCell(cell).border.left = { style: 'medium' });
    ['J55', 'J56', 'J57', 'J58', 'J59', 'J60', 'J61', 'J62', 'J63', 'J64', 'J65', 'J66', 'J67', 'J68', 'J69', 'J70', 'J71', 'J72', 'J73', 'J74', 'J75', 'J76', 'J77', 'J78', 'J79', 'J80', 'J81', 'J82', 'J83', 'J84', 'J85', 'J86', 'J87', 'J88', 'J89', 'J90', 'J91', 'J92', 'J93', 'J94', 'J95', 'J96', 'J97', 'J98', 'J99', 'J100','J101', 'J102', 'J103', 'J104', 'J105', 'J106', 'J107', 'J108', 'J109', 'J110', 'J111', 'J112'].forEach(cell => worksheet.getCell(cell).border.right = { style: 'medium' });

    worksheet.getCell('A54').border = {
        top: { style: 'medium' },
        left: { style: 'medium' },
        right: { style: 'medium' },
        bottom: { style: 'medium' }
    };


    
   // ---------------- Training resources ---------------- //

    worksheet.addRow([]); // spacer
    worksheet.mergeCells('B120:I120');
    const trainingResources = worksheet.getCell('B120');
    trainingResources.value = 'المراجع التدريبيه';
    trainingResources.font = { bold: true, size: 12 };
    trainingResources.alignment = { vertical: 'middle', horizontal: 'center' };
    trainingResources.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } };

    // Header
    worksheet.mergeCells('G121:I121');
    const c1 = worksheet.getCell('G121');
    c1.value = 'المرجع الرئيسي للمقرر';

    worksheet.mergeCells('E121:F121');
    const c2 = worksheet.getCell('E121');
    c2.value = 'المواقع الالكترونيه';

    worksheet.mergeCells('B121:D121');
    const c3 = worksheet.getCell('B121');
    c3.value = 'منصات الكترونيه';

    // Rows
    const mergedCells = [
    'G122:I122', 'E122:F122', 'B122:D122',
    'G123:I123', 'E123:F123', 'B123:D123',
    'G124:I124', 'E124:F124', 'B124:D124'
    ];

    mergedCells.forEach(range => worksheet.mergeCells(range));

    // Apply formatting to all relevant cells (alignment and font size)
    const trainingResourcesCellsToFormat = [
    'G121', 'E121', 'B121',
    'G122', 'E122', 'B122',
    'G123', 'E123', 'B123',
    'G124', 'E124', 'B124'
    ];

    trainingResourcesCellsToFormat.forEach(addr => {
    const formattedCell = worksheet.getCell(addr);
    formattedCell.alignment = { vertical: 'middle', horizontal: 'center' };
    formattedCell.font = { size: 8 };
    });

    // Borders
    const trainingResourcesTableRange = [];
    for (let i = 120; i <= 124; i++) {
    ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'].forEach(col => {
        trainingResourcesTableRange.push(`${col}${i}`);
    });
    }

    // Apply thin internal borders
    trainingResourcesTableRange.forEach(cellAddress => {
    const cell = worksheet.getCell(cellAddress);
    cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };
    });

    // Apply medium external borders
    ['B120','C120','D120','E120','F120','G120','H120','I120'].forEach(addr => worksheet.getCell(addr).border.top = { style: 'medium' });
    ['B124','C124','D124','E124','F124','G124','H124','I124'].forEach(addr => worksheet.getCell(addr).border.bottom = { style: 'medium' });
    ['B120','B121','B122','B123','B124'].forEach(addr => worksheet.getCell(addr).border.left = { style: 'medium' });
    ['I120','I121','I122','I123','I124'].forEach(addr => worksheet.getCell(addr).border.right = { style: 'medium' });


    // ---------------- Quality assessment ---------------- //
    worksheet.addRow([]); // spacer
    worksheet.mergeCells('B126:I126');
    const trainingResources2 = worksheet.getCell('B126');
    trainingResources2.value = 'تقييم جوده المقرر';
    trainingResources2.font = { bold: true, size: 12 };
    trainingResources2.alignment = { vertical: 'middle', horizontal: 'center' };
    trainingResources2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } };

    // Header
    worksheet.mergeCells('G127:I127');
    const c1_2 = worksheet.getCell('G127');
    c1_2.value = 'مجالات التقييم';

    worksheet.mergeCells('E127:F127');
    const c2_2 = worksheet.getCell('E127');
    c2_2.value = 'المقيمون';

    worksheet.mergeCells('B127:D127');
    const c3_2 = worksheet.getCell('B127');
    c3_2.value = 'طريقه التقييم';

    // Rows
    const mergedCells2 = [
    'G128:I128', 'E128:F128', 'B128:D128',
    'G129:I129', 'E129:F129', 'B129:D129',
    'G130:I130', 'E130:F130', 'B130:D130'
    ];

    mergedCells2.forEach(range => worksheet.mergeCells(range));

    // Apply formatting to all relevant cells (alignment and font size)
    const trainingResourcesCellsToFormat2 = [
    'G127', 'E127', 'B127',
    'G128', 'E128', 'B128',
    'G129', 'E129', 'B129',
    'G130', 'E130', 'B130'
    ];

    trainingResourcesCellsToFormat2.forEach(addr => {
    const formattedCell = worksheet.getCell(addr);
    formattedCell.alignment = { vertical: 'middle', horizontal: 'center' };
    formattedCell.font = { size: 8 };
    });

    // Borders
    const trainingResourcesTableRange2 = [];
    for (let i = 126; i <= 130; i++) {
    ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'].forEach(col => {
        trainingResourcesTableRange2.push(`${col}${i}`);
    });
    }

    // Apply thin internal borders
    trainingResourcesTableRange2.forEach(cellAddress => {
    const cell = worksheet.getCell(cellAddress);
    cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };
    });

    // Apply medium external borders
    ['B126','C126','D126','E126','F126','G126','H126','I126'].forEach(addr => worksheet.getCell(addr).border.top = { style: 'medium' });
    ['B130','C130','D130','E130','F130','G130','H130','I130'].forEach(addr => worksheet.getCell(addr).border.bottom = { style: 'medium' });
    ['B126','B127','B128','B129','B130'].forEach(addr => worksheet.getCell(addr).border.left = { style: 'medium' });
    ['I126','I127','I128','I129','I130'].forEach(addr => worksheet.getCell(addr).border.right = { style: 'medium' });


    // -------------------- Trainer and Head Names Table ---------------- //

    // Trainer
    worksheet.addRow([]); // spacer
    worksheet.mergeCells('I132:J132');
    const trainerName = worksheet.getCell('I132');
    trainerName.value = 'اسم المدرب';
    worksheet.getCell('G132').value = selectedTrainer.name || 'N/A';

    worksheet.mergeCells('G132:H132');
    worksheet.mergeCells('E132:F132');
    const trainerEmail = worksheet.getCell('E132');
    trainerEmail.value = 'البريد الالكتروني';

    worksheet.mergeCells('C132:D132');
    const theDate1 = worksheet.getCell('B132');
    theDate1.value = 'التاريخ';
    worksheet.getCell('A132').value = formattedDate;


    // Head
    worksheet.addRow([]); // next row
    worksheet.mergeCells('I133:J133');
    const headName = worksheet.getCell('I133');
    headName.value = 'رئيس القسم';

    // worksheet.getCell('G133').value = 'م. احمد العيسى'

    worksheet.mergeCells('G133:H133');
    worksheet.mergeCells('E133:F133');
    const headEmail = worksheet.getCell('E133');
    headEmail.value = 'البريد الالكتروني';

    worksheet.mergeCells('C133:D133');
    const theDate2 = worksheet.getCell('B133');
    theDate2.value = 'التاريخ';
    worksheet.getCell('A133').value = formattedDate;

    // Apply formatting to all relevant cells (alignment and font size)
    const trainerHeadCellsToFormat = ['A132', 'B132', 'E132', 'I132', 'A133', 'B133', 'E133', 'I133'];
    trainerHeadCellsToFormat.forEach(addr => {
    const formattedCell = worksheet.getCell(addr);
    formattedCell.alignment = { vertical: 'middle', horizontal: 'center' };
    formattedCell.font = { size: 10 };
    });

    // Borders
    const trainerHeadTableRange = [];
    for (let i = 132; i <= 133; i++) {
    ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'].forEach(col => {
        trainerHeadTableRange.push(`${col}${i}`);
    });
    }

    // Apply thin internal borders
    trainerHeadTableRange.forEach(cellAddress => {
    const cell = worksheet.getCell(cellAddress);
    cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };
    });

    // Apply medium external borders
    ['A132','B132','C132','D132','E132','F132','G132','H132','I132','J132'].forEach(addr => worksheet.getCell(addr).border.top = { style: 'medium' });
    ['A133','B133','C133','D133','E133','F133','G133','H133','I133','J133'].forEach(addr => worksheet.getCell(addr).border.bottom = { style: 'medium' });
    ['A132','A133'].forEach(addr => worksheet.getCell(addr).border.left = { style: 'medium' });
    ['J132','J133'].forEach(addr => worksheet.getCell(addr).border.right = { style: 'medium' });


        
    // -------------------------------------------------------------------------- //

    // Define widths for columns A to J
    worksheet.getColumn('A').width = 7;
    worksheet.getColumn('B').width = 7;
    worksheet.getColumn('C').width = 8;
    worksheet.getColumn('D').width = 8;
    worksheet.getColumn('E').width = 8;
    worksheet.getColumn('F').width = 7;
    worksheet.getColumn('G').width = 7;
    worksheet.getColumn('H').width = 22.5;
    worksheet.getColumn('I').width = 8;
    worksheet.getColumn('J').width = 2.5;

    // Hide all columns beyond J (starting from column 11 to 16384)
    for (let col = 11; col <= 16384; col++) {
        worksheet.getColumn(col).hidden = true; // Hide the column completely
    }



    // Generate and download the file (no conflict with 'imageBlob')
    const buffer = await workbook.xlsx.writeBuffer();
    const fileBlob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(fileBlob);
    link.download = 'Module_Template.xlsx';
    link.click();

});

// Load data on page load
loadTrainers();
loadModules();
