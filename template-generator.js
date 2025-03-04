// Function to generate the Excel template
async function generateTemplate() {
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

    // ------  Header Section -------//
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
    titleCell.value = 'Module File Template';
    titleCell.font = { size: 16, bold: true };
    titleCell.alignment = { vertical: 'middle', horizontal: 'center' };

    // --------  Trainer Information Section -------//
    worksheet.addRow([]);
    worksheet.mergeCells('B3:I3');  // Merge cells from B3 to I3
    const trainerInfoCell = worksheet.getCell('B3');
    trainerInfoCell.value = 'Trainer info';
    trainerInfoCell.font = { bold: true, size: 12 };
    trainerInfoCell.alignment = { vertical: 'middle', horizontal: 'center' };
    trainerInfoCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } };

    // Add text labels
    const cellsToFormat = ['I4', 'I5', 'I6', 'E4', 'E5', 'E6'];
    const labels = ['Trainer number', 'اسم المدرب', 'Trainer section', 'Building number', 'Office number', 'Email'];

    cellsToFormat.forEach((cellAddress, index) => {
        const cell = worksheet.getCell(cellAddress);
        cell.value = labels[index];
        cell.font = { italic: true, size: 9 };
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
    });

    // Merge cells for input fields
    worksheet.mergeCells('F4:H4');
    worksheet.mergeCells('F5:H5');
    worksheet.mergeCells('F6:H6');
    worksheet.mergeCells('B4:D4');
    worksheet.mergeCells('B5:D5');
    worksheet.mergeCells('B6:D6');


    worksheet.getRow(3).height = 20;  // Set row 3 height to 30 pixels
    worksheet.getRow(4).height = 15;  // Set row 4 height to 25 pixels
    worksheet.getRow(5).height = 15;  // Set row 5 height to 25 pixels
    worksheet.getRow(6).height = 15;  // Set row 6 height to 25 pixels


    // -------- Apply Borders -------- //
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

    // -------- Contact Section -------//

    // ------ Trainer Contact (First Instance) ----- //
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

    // ------ Head of Department Contact (Second Instance) ----- //
    worksheet.addRow([]);
    worksheet.addRow([]);
    worksheet.mergeCells('G15:I19');
    worksheet.getCell('G15').value = 'اليه التواصل مع مدرب المقرر';
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


    //  Apply Borders for Both Instances
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

    
    // --------  Module Information -------//
    worksheet.addRow([]);
    worksheet.mergeCells('B21:I21');  // Merge cells from B21 to I21
    const trainerInfoCell2 = worksheet.getCell('B21');
    trainerInfoCell2.value = 'بيانات المقرر التدريبي';
    trainerInfoCell2.font = { bold: true, size: 12 };
    trainerInfoCell2.alignment = { vertical: 'middle', horizontal: 'center' };
    trainerInfoCell2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } };

    // Add text labels
    const cellsToFormat2 = ['I22', 'I23', 'I24', 'I25', 'I26', 'E22', 'E23', 'E24', 'E25', 'E26'];
    const labels2 = ['Trainer number', 'اسم المدرب', 'Trainer section', 'Department', 'Phone number', 'Building number', 'Office number', 'Email', 'Experience', 'Address'];

    cellsToFormat2.forEach((cellAddress, index) => {
        const cell = worksheet.getCell(cellAddress);
        cell.value = labels2[index];
        cell.font = { italic: true, size: 9 };
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
    });

    // Merge cells for input fields
    worksheet.mergeCells('F22:H22');
    worksheet.mergeCells('F23:H23');
    worksheet.mergeCells('F24:H24');
    worksheet.mergeCells('F25:H25');
    worksheet.mergeCells('F26:H26');
    worksheet.mergeCells('B22:D22');
    worksheet.mergeCells('B23:D23');
    worksheet.mergeCells('B24:D24');
    worksheet.mergeCells('B25:D25');
    worksheet.mergeCells('B26:D26');

    worksheet.getRow(21).height = 20;  // Set row 21 height to 30 pixels
    worksheet.getRow(22).height = 15;  // Set row 22 height to 25 pixels
    worksheet.getRow(23).height = 15;  // Set row 23 height to 25 pixels
    worksheet.getRow(24).height = 15;  // Set row 24 height to 25 pixels
    worksheet.getRow(25).height = 15;  // Set row 25 height to 25 pixels
    worksheet.getRow(26).height = 15;  // Set row 26 height to 25 pixels

    // Module describtion
    worksheet.mergeCells('B28:I28');
    const moduleDecribtion = worksheet.getCell('B28');
    moduleDecribtion.value = 'وصف المقرر';
    moduleDecribtion.font = { bold: true, size: 12 };
    moduleDecribtion.alignment = { vertical: 'middle', horizontal: 'center' };
    moduleDecribtion.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } };

    worksheet.mergeCells('B29:I32');

    // General Goal
    worksheet.mergeCells('B34:I34');
    const generalGoal = worksheet.getCell('B34');
    generalGoal.value = 'الهدف العام';
    generalGoal.font = { bold: true, size: 12 };
    generalGoal.alignment = { vertical: 'middle', horizontal: 'center' };
    generalGoal.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } };

    worksheet.mergeCells('B35:I38');

    // ------------ Training requitments ------------ //
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

    // Safty conditions
    worksheet.mergeCells('B47:I47');  
    const safty = worksheet.getCell('B47');
    safty.value = 'تعليمات واشتراطات السلامه';
    safty.font = { bold: true, size: 12 };
    safty.alignment = { vertical: 'middle', horizontal: 'center' };
    safty.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } };

    worksheet.mergeCells('B48:I51');


    // ---------------- Training plan -------------- //
    worksheet.mergeCells('A54:J54');  
    const trainingPlan = worksheet.getCell('A54');
    trainingPlan.value = 'الخطه التدريبيه';
    trainingPlan.font = { bold: true, size: 12 };
    trainingPlan.alignment = { vertical: 'middle', horizontal: 'center' };
    trainingPlan.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } };

    // Table header
    worksheet.getCell('J55').value = 'ت'
    worksheet.mergeCells('H55:I55');
    const columnOneTitle = worksheet.getCell('H55');
    columnOneTitle.value = 'الوحدات النظريه والعمليه';
    worksheet.mergeCells('F55:G55');
    const columnTwoTitle = worksheet.getCell('F55');
    columnTwoTitle.value = 'الاهداف التفصيليه';
    worksheet.getCell('E55').value = 'الاسبوع التدريبي'
    worksheet.getCell('D55').value = 'ساعات التدريب'
    worksheet.getCell('C55').value = 'استراتيجيه التدريب'
    worksheet.getCell('B55').value = 'اليه التقييم'
    worksheet.getCell('A55').value = 'درجه التقييم'

    // Apply alignment and background color to all header cells
    const headerCells = ['J55', 'H55', 'F55', 'E55', 'D55', 'C55', 'B55', 'A55'];
    headerCells.forEach(cellAddress => {
        const cell = worksheet.getCell(cellAddress);
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD3D3D3' } }; // Grey background
        cell.font = { size: 8 };
    });

   // Create 40 empty rows with sequence numbers in column J
    for (let i = 1; i <= 40; i++) {
        const rowNumber = 55 + i;
        const cell = worksheet.getCell(`J${rowNumber}`);
        cell.value = i; // Fill sequence numbers in column J
        cell.font = { size: 8 };
        worksheet.mergeCells(`H${rowNumber}:I${rowNumber}`); // Merge H and I in each row
        worksheet.mergeCells(`F${rowNumber}:G${rowNumber}`); // Merge F and G in each row
    }

    // Apply borders
    const tableRange = [];
    for (let i = 55; i <= 95; i++) {
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
    ['A95', 'B95', 'C95', 'D95', 'E95', 'F95', 'G95', 'H95', 'I95', 'J95'].forEach(cell => worksheet.getCell(cell).border.bottom = { style: 'medium' });
    ['A55', 'A56', 'A57', 'A58', 'A59', 'A60', 'A61', 'A62', 'A63', 'A64', 'A65', 'A66', 'A67', 'A68', 'A69', 'A70', 'A71', 'A72', 'A73', 'A74', 'A75', 'A76', 'A77', 'A78', 'A79', 'A80', 'A81', 'A82', 'A83', 'A84', 'A85', 'A86', 'A87', 'A88', 'A89', 'A90', 'A91', 'A92', 'A93', 'A94', 'A95'].forEach(cell => worksheet.getCell(cell).border.left = { style: 'medium' });
    ['J55', 'J56', 'J57', 'J58', 'J59', 'J60', 'J61', 'J62', 'J63', 'J64', 'J65', 'J66', 'J67', 'J68', 'J69', 'J70', 'J71', 'J72', 'J73', 'J74', 'J75', 'J76', 'J77', 'J78', 'J79', 'J80', 'J81', 'J82', 'J83', 'J84', 'J85', 'J86', 'J87', 'J88', 'J89', 'J90', 'J91', 'J92', 'J93', 'J94', 'J95'].forEach(cell => worksheet.getCell(cell).border.right = { style: 'medium' });



    // ----------- Training resources -------- //
    worksheet.addRow([]);
    worksheet.mergeCells('B97:I97');  
    const trainingResources = worksheet.getCell('B97');
    trainingResources.value = 'المراجع التدريبيه';
    trainingResources.font = { bold: true, size: 12 };
    trainingResources.alignment = { vertical: 'middle', horizontal: 'center' };
    trainingResources.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } };

    // Header
    worksheet.mergeCells('G98:I98');
    const c1 = worksheet.getCell('G98');
    c1.value = 'المرجع الرئيسي للمقرر';

    worksheet.mergeCells('E98:F98');
    const c2 = worksheet.getCell('E98');
    c2.value = 'المواقع الالكترونيه';

    worksheet.mergeCells('B98:D98');
    const c3 = worksheet.getCell('B98');
    c3.value = 'منصات الكترونيه';

    // Rows
    const mergedCells = ['G99:I99', 'E99:F99', 'B99:D99', 
                        'G100:I100', 'E100:F100', 'B100:D100', 
                        'G101:I101', 'E101:F101', 'B101:D101'];

    mergedCells.forEach(range => worksheet.mergeCells(range));

    // Apply formatting to all relevant cells (alignment and font size)
    const trainingResourcesCellsToFormat = ['G98', 'E98', 'B98', 'G99', 'E99', 'B99', 'G100', 'E100', 'B100', 'G101', 'E101', 'B101'];
    trainingResourcesCellsToFormat.forEach(cell => {
        const formattedCell = worksheet.getCell(cell);
        formattedCell.alignment = { vertical: 'middle', horizontal: 'center' };
        formattedCell.font = { size: 8 };
    });

    // Apply borders
    const trainingResourcesTableRange = [];
    for (let i = 97; i <= 101; i++) {
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
    ['B97', 'C97', 'D97', 'E97', 'F97', 'G97', 'H97', 'I97'].forEach(cell => worksheet.getCell(cell).border.top = { style: 'medium' });
    ['B101', 'C101', 'D101', 'E101', 'F101', 'G101', 'H101', 'I101'].forEach(cell => worksheet.getCell(cell).border.bottom = { style: 'medium' });
    ['B97', 'B98', 'B99', 'B100', 'B101'].forEach(cell => worksheet.getCell(cell).border.left = { style: 'medium' });
    ['I97', 'I98', 'I99', 'I100', 'I101'].forEach(cell => worksheet.getCell(cell).border.right = { style: 'medium' });

    // -------------- Quality assessment --------------- //
    worksheet.addRow([]);
    worksheet.mergeCells('B103:I103');  
    const trainingResources2 = worksheet.getCell('B103');
    trainingResources2.value = 'تقييم جوده المقرر';
    trainingResources2.font = { bold: true, size: 12 };
    trainingResources2.alignment = { vertical: 'middle', horizontal: 'center' };
    trainingResources2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } };

    // Header
    worksheet.mergeCells('G104:I104');
    const c1_2 = worksheet.getCell('G104');
    c1_2.value = 'مجالات التقييم';

    worksheet.mergeCells('E104:F104');
    const c2_2 = worksheet.getCell('E104');
    c2_2.value = 'المقيمون';

    worksheet.mergeCells('B104:D104');
    const c3_2 = worksheet.getCell('B104');
    c3_2.value = 'طريقه التقييم';

    // Rows
    const mergedCells2 = ['G105:I105', 'E105:F105', 'B105:D105', 
                        'G106:I106', 'E106:F106', 'B106:D106', 
                        'G107:I107', 'E107:F107', 'B107:D107'];

    mergedCells2.forEach(range => worksheet.mergeCells(range));

    // Apply formatting to all relevant cells (alignment and font size)
    const trainingResourcesCellsToFormat2 = ['G104', 'E104', 'B104', 'G105', 'E105', 'B105', 'G106', 'E106', 'B106', 'G107', 'E107', 'B107'];
    trainingResourcesCellsToFormat2.forEach(cell => {
        const formattedCell = worksheet.getCell(cell);
        formattedCell.alignment = { vertical: 'middle', horizontal: 'center' };
        formattedCell.font = { size: 8 };
    });

    // Apply borders
    const trainingResourcesTableRange2 = [];
    for (let i = 103; i <= 107; i++) {
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
    ['B103', 'C103', 'D103', 'E103', 'F103', 'G103', 'H103', 'I103'].forEach(cell => worksheet.getCell(cell).border.top = { style: 'medium' });
    ['B107', 'C107', 'D107', 'E107', 'F107', 'G107', 'H107', 'I107'].forEach(cell => worksheet.getCell(cell).border.bottom = { style: 'medium' });
    ['B103', 'B104', 'B105', 'B106', 'B107'].forEach(cell => worksheet.getCell(cell).border.left = { style: 'medium' });
    ['I103', 'I104', 'I105', 'I106', 'I107'].forEach(cell => worksheet.getCell(cell).border.right = { style: 'medium' });


    // -------------------- Trainer and Head Names Table ------------ // 

    // Trainer
    worksheet.addRow([]);
    worksheet.mergeCells('I109:J109');
    const trainerName = worksheet.getCell('I109');
    trainerName.value = 'اسم المدرب';

    worksheet.mergeCells('G109:H109');
    worksheet.mergeCells('E109:F109');
    const trainerEmail = worksheet.getCell('E109');
    trainerEmail.value = 'البريد الالكتروني';

    worksheet.mergeCells('C109:D109');
    const theDate1 = worksheet.getCell('B109');
    theDate1.value = 'التاريخ';


    // Head
    worksheet.addRow([]);
    worksheet.mergeCells('I110:J110');
    const headName = worksheet.getCell('I110');
    headName.value = 'رئيس القسم';

    worksheet.mergeCells('G110:H110');
    worksheet.mergeCells('E110:F110');
    const headEmail = worksheet.getCell('E110');
    headEmail.value = 'البريد الالكتروني';

    worksheet.mergeCells('C110:D110');
    const theDate2 = worksheet.getCell('B110');
    theDate2.value = 'التاريخ';

    // Apply formatting to all relevant cells (alignment and font size)
    const trainerHeadCellsToFormat = ['A109', 'B109', 'E109', 'I109', 'A110', 'B110', 'E110', 'I110'];
    trainerHeadCellsToFormat.forEach(cell => {
        const formattedCell = worksheet.getCell(cell);
        formattedCell.alignment = { vertical: 'middle', horizontal: 'center' };
        formattedCell.font = { size: 10 };
    });

    // Apply borders
    const trainerHeadTableRange = [];
    for (let i = 109; i <= 110; i++) {
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
    ['A109', 'B109', 'C109', 'D109', 'E109', 'F109', 'G109', 'H109', 'I109', 'J109'].forEach(cell => worksheet.getCell(cell).border.top = { style: 'medium' });
    ['A110', 'B110', 'C110', 'D110', 'E110', 'F110', 'G110', 'H110', 'I110', 'J110'].forEach(cell => worksheet.getCell(cell).border.bottom = { style: 'medium' });
    ['A109', 'A110'].forEach(cell => worksheet.getCell(cell).border.left = { style: 'medium' });
    ['J109', 'J110'].forEach(cell => worksheet.getCell(cell).border.right = { style: 'medium' });



    
    // -------------------------------------------------------------------------- //

    // Adjust column widths
    worksheet.columns.forEach(column => {
        column.width = 10;
    });

    // Set default column width to 100 pixels
    ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'].forEach(col => {
        worksheet.getColumn(col).width = 8.5;  // Set all used columns to a narrower width
    });
    

    // Generate and download the file (no conflict with 'imageBlob')
    const buffer = await workbook.xlsx.writeBuffer();
    const fileBlob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(fileBlob);
    link.download = 'Module_Template.xlsx';
    link.click();
}

// Trigger the function on button click
document.getElementById('generateTemplateButton').addEventListener('click', generateTemplate);
