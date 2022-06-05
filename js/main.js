var index = 2;//初始下标
var indexArr= new Array();
indexArr.push(1);
// indexArr.push(2);
// indexArr.push(3);

function addRow() {
    index++;
    indexArr.push(index);
    var showHtml = $("#travel_info_").html();
    var html = '<div class="travel_info" id="travel_info'+index +'">'+showHtml+"</div>";
    html = html.replace(/##/g,index);
    // $("#info").before($(html));
    // document.getElementById("Info").appendChild($(html));
    $(html).appendTo($($("#info")));
    console.log(html);
}

function deleteRow(inde){
    $("#travel_info" + inde).remove();
    var a = indexArr.indexOf(parseInt(inde));
    if (a > -1) {
        indexArr.splice(a, 1);
        console.log("当前下标数组",indexArr);
    }
}


 function json2Sheet () {
    let json = [
        {
            "姓名": "张三",
            "性别": "男",
            "年龄": 18
        },
        {
            "姓名": "李四",
            "性别": "女",
            "年龄": 19
        },
        {
            "姓名": "王二麻",
            "性别": "未知",
            "年龄": 20
        }
    ]

    // 实例化一个工作簿
    let book = XLSX.utils.book_new()

    // 实例化一个Sheet
    let sheet = XLSX.utils.json_to_sheet(json, {
        header: ['姓名', '性别', '年龄']
    })

    // 将Sheet写入工作簿
    XLSX.utils.book_append_sheet(book, sheet, 'Sheet1')

    // 写入文件，直接触发浏览器的下载
    XLSX.writeFile(book, 'json2Sheet.xlsx')
}


function array2Sheet () {
    let data = [
        ['姓名', '性别', '年龄'],
        ['张三', '男', '18'],
        ['李四', '女', '19'],
        ['王二麻', '未知', '20']
    ]

    // 实例化一个工作簿
    let book = XLSX.utils.book_new()

    // 实例化一个Sheet
    let sheet = XLSX.utils.aoa_to_sheet(data)

    // 将Sheet写入工作簿
    XLSX.utils.book_append_sheet(book, sheet, 'Sheet1')

    // 写入文件，直接触发浏览器的下载
    XLSX.writeFile(book, 'array2Sheet.xlsx')
}


function parseExcel (fileDom) {
    let file = fileDom.files[0]
    let reader = new FileReader()
    let rABS = typeof FileReader !== "undefined" && (FileReader.prototype||{}).readAsBinaryString
    if (rABS) {
        reader.readAsBinaryString(file)
    } else {
        reader.readAsArrayBuffer(file)
    }
    reader.onload = function(e) {
        let data = e.target.result
        if (!rABS) {
            data = new Uint8Array(data)
        }
        let workBook = XLSX.read(data, {type: rABS ? 'binary' : 'array'})
        workBook.SheetNames.forEach(name => {

            let sheet = workBook.Sheets[name]
            let json = XLSX.utils.sheet_to_json(sheet, {
                raw: false,
                header: 1
            })

            console.log(json);
            let book = XLSX.utils.book_new();

            // 实例化一个Sheet
            let sheet1 = XLSX.utils.json_to_sheet(json);
        
            // 将Sheet写入工作簿
            XLSX.utils.book_append_sheet(book, sheet, 'Sheet1');
        
            // 写入文件，直接触发浏览器的下载
            XLSX.writeFile(book, 'json2Sheet.xlsx');       
        })
    }
}

let titleCellStyle = {
    font: {name:"宋体",sz: 13,bold: true,color: { rgb: '000000' }},
    alignment: {horizontal: 'center',vertical: 'center',wrap_text: 'false'},
    border: {top: {style: "thin",},bottom: {style: "thin",},left: {style: "thin",},right: {style: "thin"}}
};

let defaultCellStyle = {
    font: {name:"宋体",sz: 11,bold: false,color: { rgb: '000000' }},
    alignment: {horizontal: 'center',vertical: 'center',wrap_text: 'false'},
    border: {top: {style: "thin",},bottom: {style: "thin",},left: {style: "thin",},right: {style: "thin"}}
};

descCellStyle = {
    font: {name:"宋体",sz: 11,bold: false,color: { rgb: '000000' }},
    alignment: {horizontal: 'bottom',vertical: 'center',wrap_text: 'false'},
    border: {top: {style: "thin",},bottom: {style: "thin",},left: {style: "thin",},right: {style: "thin"}}
};

headerCellStyle = {
    font: {name:"宋体",sz: 11,bold: true,color: { rgb: '000000' }},
    alignment: {horizontal: 'center',vertical: 'center',wrap_text: 'false'},
    border: {top: {style: "thin",},bottom: {style: "thin",},left: {style: "thin",},right: {style: "thin"}}
};

secondCellStyle = {
    font: {name:"宋体",sz: 11,bold: true,color: { rgb: '000000' }},
    alignment: {horizontal: 'bottom',vertical: 'center',wrap_text: 'false'},
    border: {top: {style: "thin",},bottom: {style: "thin",},left: {style: "thin",},right: {style: "thin"}}
};

lastCellStyle = {
    font: {name:"宋体",sz: 11,bold: false,color: { rgb: '000000' }},
    alignment: {horizontal:'top',vertical: 'center',wrap_text: 'true'},
    border: {top: {style: "thin",},bottom: {style: "thin",},left: {style: "thin",},right: {style: "thin"}}
};


function openDownloadDialog(url, saveName) {
    if (typeof url == 'object' && url instanceof Blob) {
        url = URL.createObjectURL(url); // 创建blob地址
    }
    var aLink = document.createElement('a');
    aLink.href = url;
    aLink.download = saveName || ''; // HTML5新增的属性，指定保存文件名，可以不要后缀，注意，file:///模式下不会生效
    var event;
    if (window.MouseEvent) event = new MouseEvent('click');
    else {
        event = document.createEvent('MouseEvents');
        event.initMouseEvent('click', true, false, window, 0, 0, 0, 0, 0, false, false, false, false, 0,
            null);
    }
    aLink.dispatchEvent(event);
}


function sheet_from_array_of_arrays(data) {
    const ws = {};
    const range = {
      s: {
        c: 10000000,
        r: 10000000
      },
      e: {
        c: 0,
        r: 0
      }
    };
    for (let R = 0; R !== data.length; ++R) {
      for (let C = 0; C !== data[R].length; ++C) {
        if (range.s.r > R) range.s.r = R;
        if (range.s.c > C) range.s.c = C;
        if (range.e.r < R) range.e.r = R;
        if (range.e.c < C) range.e.c = C;
        let cell={};
        if(R==0 && C==0){
            cell = {
                v: data[R][C],
                s: titleCellStyle
            };
        }else if(R==1){
            cell = {
                v: data[R][C],
                s: descCellStyle
            };
        }else if(R==2){
            cell = {
                v: data[R][C],
                s: headerCellStyle
            };
        }else if(R==data.length-2){
            cell = {
                v: data[R][C],
                s: secondCellStyle
            };
        }else if(R==data.length-1 && C==0){
            cell = {
                v: data[R][C],
                s: lastCellStyle
            };
        }else{
            cell = {
                v: data[R][C],
                s: defaultCellStyle
            };
        } 

        if (cell.v == null) {
            continue;
        }
        const cell_ref = XLSX.utils.encode_cell({
          c: C,
          r: R
        });
        /* TEST: proper cell types and value handling */
        if (typeof cell.v === 'number') {
          cell.t = 'n';
        } else if (typeof cell.v === 'boolean') {
          cell.t = 'b';
        } else if (cell.v instanceof Date) {
          cell.t = 'n';
          cell.z = XLSX.SSF._table[14];
          cell.v = this.dateNum(cell.v);
        } else {
          cell.t = 's';
        }
        ws[cell_ref] = cell;
      }
    }
    /* TEST: proper range */
    if (range.s.c < 10000000) {
        ws['!ref'] = XLSX.utils.encode_range(range);
    }
    return ws;
}


function sheet2blob(sheet){
    sheetName = "sheet1";
    const workbook = {SheetNames:[sheetName],Sheets:{}};
    workbook.Sheets[sheetName]=sheet;
    const wopts ={bookType:"xlsx",bookSST:false,type:'binary'};
    const wbout= XLSX.write(workbook,wopts,{defaultCellStyle:defaultCellStyle});
    const blob =new Blob([s2ab(wbout)],{type:"application/octet-stream"});

    function s2ab(s){
        const buf = new ArrayBuffer(s.length);
        const view = new Uint8Array(buf);
        for (let i=0;i!=s.length;i++){
            view[i]=s.charCodeAt(i)&0xFF;
        }
        return buf;
    }
    return blob;
}

function submit(){
    let all_data = [];

    let heji = [];

    var total_oil=0;
    var total_rent=0;
    var total_lufei=0;
    var total_tingche=0;
    var total_chaoshi=0;
    var total_chaokil=0;
    var total_food=0;
    var total_hotel=0;
    var total_final=0;
    
    var selector1  = document.getElementById("company");
    var company_index=selector1.selectedIndex;
    var company = selector1.options[company_index].text;

    var selector2  = document.getElementById("car_type");
    var car_type_index=selector2.selectedIndex;
    var car_type = selector2.options[car_type_index].text;

    var car_number_temp  = document.getElementById("car_number").value;
    console.log(car_number_temp);
    var car_number = car_number_temp.substring(0,8);
    console.log(car_number);

    var person  = document.getElementById("person").value;
    console.log(person);

    
    var date = new Date();
    var year = date.getFullYear();
    title = company+"\n"+year+"年用车清单";
    description = "用车单位："+company+"    "+person;
    console.log(description);


    console.log(title);
    temp_data = [];
    temp_data.push(title);
    for(var t=0;t<14;t++){
        temp_data.push("");
    }
    all_data.push(temp_data);

    temp_data1 = [];
    temp_data1.push(description);
    for(var t=0;t<14;t++){
        temp_data1.push("");
    }
    all_data.push(temp_data1);
    
    console.log(car_type);

    table_header = ["序号","日期","车型","车牌","用车时间","基本行程","租金","油费","路费","停车费","超时","超公里","餐费","住宿费","合计"];
    all_data.push(table_header);
    console.log(car_number);
    
    var num_of_null =0

    for (let i = 0; i < indexArr.length; i++) {
        var info_temp = "#travel_info"+indexArr[i];
        console.log(info_temp);
        var data = [];
        data.push(i+1);
        var Input = $(info_temp).find("input");
        
        if(Input[0].value==""){
            num_of_null = num_of_null+1;
            continue;
        }
        var total = 0;
        console.log(Input);
        
        for (var j = 0;j<Input.length;j++){
            // console.log("data: ",Input[j].value)

            if(i==0){
                if(j==0){
                    let riqi_list = Input[j].value.split("-");
                    console.log(riqi_list);
                    
                    var month;
                    if (riqi_list[1][0]=="0"){
                        month = riqi_list[1][1];
                    }else{
                        month = riqi_list[1];
                    }

                    var day;

                    if (riqi_list[2][0]=="0"){
                        day = riqi_list[2][1];
                    }else{
                        day = riqi_list[2];
                    }

                    var riqi = month+"月"+day+"日";
                    data.push(riqi);
                }else if(j==1){
                    data.push(car_type);
                    data.push(car_number);
                    var start_time = Input[j].value;
                    var end_time = Input[j+1].value; 
                    data.push(start_time+"-"+end_time);
                    
                }else if(j==2){
                    continue;
                }else{
                    data.push(Input[j].value);
                }
                if(j>=4){
                    total = total + Number(Input[j].value);
                    if(j==4){
                        total_rent=total_rent+Number(Input[j].value);
                    }
                    if(j==5){
                        total_oil=total_oil+Number(Input[j].value);
                    }
                    if(j==6){
                        total_lufei=total_lufei+Number(Input[j].value);
                    }
                    if(j==7){
                        total_tingche=total_tingche+Number(Input[j].value);
                    }
                    if(j==8){
                        total_chaoshi=total_chaoshi+Number(Input[j].value);
                    }
                    if(j==9){
                        total_chaokil=total_chaokil+Number(Input[j].value);
                    }
                    if(j==10){
                        total_food=total_food+Number(Input[j].value);
                    }
                    if(j==11){
                        total_hotel=total_hotel+Number(Input[j].value);
                    }
                }
            }else{
                if(j==0){
                    let riqi_list = Input[j].value.split("-");
                    var month;
                    if (riqi_list[1][0]=="0"){
                        month = riqi_list[1][1];
                    }else{
                        month = riqi_list[1];
                    }

                    var day;
                    if (riqi_list[2][0]=="0"){
                        day = riqi_list[2][1];
                    }else{
                        day = riqi_list[2];
                    }

                    var riqi = month+"月"+day+"日";
                    // var riqi = riqi_list[0]+"/"+riqi_list[1]+"/"+riqi_list[2];
                    data.push(riqi);
                }else if(j==1){
                    data.push("");
                    data.push("");
                    var start_time = Input[j].value;
                    var end_time = Input[j+1].value; 
                    data.push(start_time+"-"+end_time);
                    
                }else if(j==2){
                    continue;
                }else{
                        data.push(Input[j].value);
                }
                
                if(j>=4){
                    total = total + Number(Input[j].value);

                    if(j==4){
                        total_rent=total_rent+Number(Input[j].value);
                    }
                    if(j==5){
                        total_oil=total_oil+Number(Input[j].value);
                    }
                    if(j==6){
                        total_lufei=total_lufei+Number(Input[j].value);
                    }
                    if(j==7){
                        total_tingche=total_tingche+Number(Input[j].value);
                    }
                    if(j==8){
                        total_chaoshi=total_chaoshi+Number(Input[j].value);
                    }
                    if(j==9){
                        total_chaokil=total_chaokil+Number(Input[j].value);
                    }
                    if(j==10){
                        total_food=total_food+Number(Input[j].value);
                    }
                    if(j==11){
                        total_hotel=total_hotel+Number(Input[j].value);
                    }
                    
                }
            }
        }
        console.log("total: ",total_rent);
        total_final = total_final + total;
        data.push(total);

        all_data.push(data);
    }
    all_data.push(["","","","","","","","","","","","","","",""]);

    var heji_list=["合计","","","","",""];


    heji_list.push(total_rent);
    heji_list.push(total_oil);
    heji_list.push(total_lufei);
    heji_list.push(total_tingche);
    heji_list.push(total_chaoshi);
    heji_list.push(total_chaokil);
    heji_list.push(total_food);
    heji_list.push(total_hotel);
    heji_list.push(total_final);
    all_data.push(heji_list);

    var ps ="备注：账户名称：广州广源汽车租赁有限公司    银行账号：1209-1816-4310-902    开户行：招商银行淘金支行，需转费用  "+total_final+"  元。\n●请贵单位转账时备注用车单据号（052224）"
    
    all_data.push([ps,"","","","","","","","","","","","","",""]);
    
    all_data.push(["广州广源汽车租赁有限公司（财务部）","","","","","","","","","","","","","",""]);

    console.log(all_data);

    let book = XLSX.utils.book_new();

    // 实例化一个Sheet
    let sheet = sheet_from_array_of_arrays(all_data);

    
    const merge = [{s:{r:1,c:0},e:{r:1,c:14}},
                   {s:{r:0,c:0},e:{r:0,c:14}},
                   {s:{r:3,c:2},e:{r:2+indexArr.length-num_of_null,c:2}},
                   {s:{r:3,c:3},e:{r:2+indexArr.length-num_of_null,c:3}},
                   {s:{r:3+indexArr.length-num_of_null,c:0},e:{r:3+indexArr.length-num_of_null,c:14}},
                   {s:{r:4+indexArr.length-num_of_null,c:0},e:{r:4+indexArr.length-num_of_null,c:5}},
                   {s:{r:5+indexArr.length-num_of_null,c:0},e:{r:5+indexArr.length-num_of_null,c:14}},
                   {s:{r:6+indexArr.length-num_of_null,c:0},e:{r:6+indexArr.length-num_of_null,c:14}},
                ];
    
    sheet["!merges"] = merge;

    sheet['!cols'] = [{ wpx: 44 }, { wpx: 68 }, { wpx: 68 }, { wpx: 83 }, { wpx: 75 },{ wpx: 102 }, { wpx: 50 }, { wpx: 50 }, { wpx: 50 }, { wpx: 50 },{ wpx: 50 }, { wpx: 50 }, { wpx: 50 }, { wpx: 50 }, { wpx: 50 }];
    sheet['!rows']=[];
    for(var m=0;m<4+indexArr.length;m++){
        if(m==1){
            sheet['!rows'].push({ hpx: 26});
        }else{
            sheet['!rows'].push({ hpx: 35});
        }
    }
    sheet['!rows'].push({ hpx: 35});
    sheet['!rows'].push({ hpx: 45});
    sheet['!rows'].push({ hpx: 35});

    //最终调整格式
    

    console.log(sheet);

    // 将Sheet写入工作簿
    XLSX.utils.book_append_sheet(book, sheet, 'Sheet1');

    // 写入文件，直接触发浏览器的下载
    XLSX.writeFile(book, 'array2Sheet.xlsx');

}





