const axios = require('axios');
const cheerio = require('cheerio');
const iconv = require('iconv-lite');
const Excel = require('exceljs');

// excel의 column생성을 위한 array
const columns = new Array();
// excel의 row를 채우기 위한 array
const rows = new Array();

const DOMAIN = 'https://finance.naver.com';
const URL = '/sise/sise_market_sum.naver?sosok=1&page=';
let pageNo = 1;
let lastPage;

async function main(){

    let tempColumn = new Array();
    let tempDeatilColumn = new Array();

    // 전체 순위 정보 크롤링[S]
    do{
        const decoded = await croller(URL+pageNo);
        const $ = cheerio.load(decoded);

        // step1. 첫 페이지를 긁어서 페이지 번호와 첫페이지 정보를 세팅한다
        // step2. 회사 정보가 있는 table 크롤링
        const kospiTable = $('table.type_2');
        const kospiHeader = kospiTable.find('thead th');
        const kospiBody = kospiTable.find('tbody tr');

        // step4. column은 첫번째 페이지 긁을때에만 생성
        if(pageNo == 1){
            const navigation = $('table.Nnavi');
            lastPage = navigation.find('td.pgRR a').prop('href').split('page=')[1];
    
            // step5. 첫페이지라면 페이지 하단 네비게이션 바를 긁어서 마지막 페이지를 찾는다
            kospiHeader.each((idx, el)=> {
                let columnObj = new Object();
                columnObj.header = $(el).text().replaceAll('\n', '').replaceAll('\t', '');
                columnObj.key = $(el).text().replaceAll('\n', '').replaceAll('\t', '');
                columns.push(columnObj);
                tempColumn.push(columnObj);
            })
        
            // 상세페이지 column생성
            for(i=0; i< 11; i++){
                let columnObj = new Object();
                let keyName = (i < 4)? 'st' : 'Q';
                let keySeq;

                if(i <10 ){
                    if(i >= 4)  keySeq = i-3;
                    else        keySeq = i+1;

                    columnObj.header = String(keySeq)+keyName;
                    columnObj.key = String(keySeq)+keyName;
                    
                }else{
                    columnObj.header = '비고';
                    columnObj.key = '비고';
                }

                columns.push(columnObj);    // 실제 액셀 그릴때 쓸놈
                tempDeatilColumn.push(columnObj); // 데이터 생성할때 쓸놈
            }
        }
        // step6. 회사정보들을 obj array로 생성
        kospiBody.each( async (idx, el) => {
        
            // 미관상 탭을 나누기 위해 5개? 마다 blank_08 blank_06같은 애들이 있음
            // blank class가없는 td가 첫 td면 이 tr은 실제 우리가 원하는 값이 있는 row
            if($(el).find('td').length > 1){
                let companyList = $(el).find('td');
                let rowObj = new Object();
                let hrefTarget;
                companyList.each((idx, item) => {
                    // 뭔....뭔 긁어오기 귀찮게 자꾸 해놨다...하...그래도 트라이
                    
                    // a태그에 href속성이 있으면 걔는 회사별 상세페이지가 있는 tr  
                    if($(item).find('a').length >= 1 && !$(item).hasClass('center')){
                        hrefTarget = $(item).find('a').attr('href');
                    }

                    rowObj[tempColumn[idx].key] = $(item).text().replaceAll('\n', '').replaceAll('\t', '');

                })

                let price = Number(rowObj['시가총액'].replaceAll(',',''));
                console.log(price)

                // 각 회사별 상세페이지 크롤링[S]
                // 시총이 350이상 800사이일 때 상세페이지 조회  
                if( 350 <= price && price <= 800 ){

                    const detailDecoded = await croller(hrefTarget)
                    const detail$ = cheerio.load(detailDecoded);

                    const copAnalysis = detail$('div.cop_analysis table');

                    // 회사별 데이터가 몇년도거인지 그냥 스트링으로라도 가져올수있을라나
                    let tempStr = "";
                    let yCnt = 0;
                    let qCnt = 0;
                    tempStr+='##연간## ';

                    let tableHeader = copAnalysis.find('thead tr')[1];
                    detail$(tableHeader).find('th').each((idx, item)=>{
                        tempStr += '  '+detail$(item).text().replaceAll('\n', '').replaceAll('\t', '');
                        
                        if(detail$(item).hasClass('t_line')){
                            tempStr += ' ##분기## ';
                        }
                    })

                    // 회사별 영업이익 get  th_cop_anal9
                    let profitTable = copAnalysis.find('tbody tr');

                    profitTable.each((idx, item)=>{
                        // tr 내 th에 th_cop_anal9 클래스를 가진 tr이 영업이익 row
                        if(detail$(item).find('th').hasClass('th_cop_anal9')){
                            let profitData = detail$(item).find('td');
                            detail$(profitData).each((idx, el) => {
                                // 어우 일단 E붙은애도 같이 삽입
                                // if(!$(el).hasClass('cell_strong')){
                                // }
                                rowObj[tempDeatilColumn[idx].key] = $(el).text().replaceAll('\n', '').replaceAll('\t', '');
                                
                            })
                        }
                    })
                    
                    rowObj['비고'] = tempStr;
                    console.log(rowObj)

                    rows.push(rowObj);
                }
                // 각 회사별 상세페이지 크롤링[E]
            }
        
        });

        pageNo ++;
    }
    while(pageNo <= lastPage);
    // 전체 순위 정보 크롤링[E]

    await generateToExcel(columns, rows);
}




/**
 * url의 html을 가져와서 euc-kr로 변경 후 return
 * @param {String} url
 * @returns 
 */
async function croller(_url){
    // step1. HTML 로드
    let html = await axios.get(
        DOMAIN+_url
     , {
        responseType: 'arraybuffer',
     });

    // 이건 왜 euc-kr로 긁어와질까... 페이지는 utf-8같은데...희한하다..
    // step2. 긁어온 HTML 인코딩 변경
    let contentType = html.headers['content-type'];
    let charset = contentType.includes('charset=') ? contentType.split('charset=')[1] : 'UTF-8';
    
    // iconv를 사용해서 axios로 받아온 data를 decode
    const decoded = iconv.decode(html.data, charset);

    return decoded;
}

async function generateToExcel(columns, rows){
    //엑셀 워크북 생성 및 시트 생성
    let today = new Date();
    let year = today.getFullYear();
    let month = today.getMonth()+1 
    let date = today.getDate();

    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet('sheet1');

    // sheet.mergeCells('A1:C1');

    //대표행(타이틀행) 설정 및 입력
    worksheet.columns = columns;

    rows.map( (item, index) => {
        worksheet.addRow(item);
      });

    
    //엑셀 데이터 저장
    await workbook.xlsx.writeFile(`${year}${(month < 10) ? '0'+month : month}${(date < 10) ? '0'+date : date}${today.getHours()}${today.getMinutes()}${today.getSeconds()}.xlsx`);

    console.log('종료')
}


main();