const xl = require('excel4node');
const moment = require('moment');

let style = [
    {
      font: {
        color: 'black',
        size: 12,
        bold: true,
      },
      alignment: {
        wrapText: true,
        horizontal: 'center',
        vertical: 'center'
      },
      fill: {
        type: 'pattern',
        patternType: 'solid',
        bgColor: '#70AD47',
        fgColor: '#70AD47',
      },
      border: {
        left: {
          style: 'thin',
          color: 'black',
        },
        right: {
          style: 'thin',
          color: 'black',
        },
        top: {
          style: 'thin',
          color: 'black',
        },
        bottom: {
          style: 'thin',
          color: 'black',
        },
        outline: false,
      },
    },
    {
      font: {
        color: '#2c3e50',
        size: 12,
      },
      alignment: {
        wrapText: true,
        horizontal: 'center',
        vertical: 'center'
      },
      border: {
        left: {
          style: 'thin',
          color: 'black',
        },
        right: {
          style: 'thin',
          color: 'black',
        },
        top: {
          style: 'thin',
          color: 'black',
        },
        bottom: {
          style: 'thin',
          color: 'black',
        },
        outline: false,
      },
    }
];


const createExcel = async (title, body, style, sheet) => {
    const wb = new xl.Workbook();
    const ws = wb.addWorksheet(sheet);
    let grabStyle = [];



    style.map(v => {
      grabStyle.push(wb.createStyle(v));
    });


    // ws.cell(1,1,0,0,false)
    // .string(body[0].title)
    // .style(grabStyle[1])
  
    title.map(v => {
      (typeof v.title != 'number') ? ws.cell(v.cell.a, v.cell.b, v.cell.c, v.cell.d, v.cell.e).string(`${v.title}`).style(grabStyle[v.style])
        : ws.cell(v.cell.a, v.cell.b, v.cell.c, v.cell.d, v.cell.e).number(v.title).style(grabStyle[v.style]);
    });
  
    body.map(v => {
      (typeof v.title != 'number') ? ws.cell(v.cell.a, v.cell.b, v.cell.c, v.cell.d, v.cell.e).string(`${v.title}`).style(grabStyle[v.style])
        : ws.cell(v.cell.a, v.cell.b, v.cell.c, v.cell.d, v.cell.e).number(v.title).style(grabStyle[v.style]);
    });

    ws.column(1).setWidth(40)
    ws.column(2).setWidth(30)
    // ws.column(3).setWidth(30)
    // ws.column(4).setWidth(30)
  

    wb.write('test_excel.xlsx');
}



(async () => {

    // start customer process
    let title = [
        {
          cell: { a: 2, b: 3, c: 0, d: 4, e: true },
          title: 'Status',
          style: 0
        },
        {
            cell: { a: 2, b: 1, c: 4, d: 0, e: true },
            title: 'Customer',
            style: 0
        },
        {
            cell: { a: 2, b: 2, c: 4, d: 0, e: true },
            title: 'Transaction Total',
            style: 0
        },
        {
            cell: { a: 3, b: 3, c: 4, d: 3, e: true },
            title: 'Success',
            style: 0
        },
        {
            cell: { a: 3, b: 4, c: 4, d: 4, e: true },
            title: 'Failed',
            style: 0
        }
      ];

      let body = []


      formatStartDate = moment().format('DD MMMM YYYY');
      formatEndDate = moment().format('DD MMMM YYYY');

      const customerSampleData = [
          {
              customerName: 'Jasa Marga',
              transcationTotal: 1,
              success: 1,
              failed: 0
          },
          {
            customerName: 'Jasa Marga 1',
            transcationTotal: 10,
            success: 5,
            failed: 5
        },
        {
            customerName: 'Jasa Marga 2',
            transcationTotal: 20,
            success: 10,
            failed: 10
        },
        {
            customerName: 'Makira 2',
            transcationTotal: 30,
            success: 10,
            failed: 10
        }
      ]

      let allTransaction = 0;
      let totalSuccess = 0;
      let totalFailed = 0;
      const startRowCustomer = 4;

      customerSampleData.map((data, i) => {
          const newIndex = i+1;

          allTransaction+=data.transcationTotal;
          totalSuccess+=data.success;
          totalFailed+=data.failed;

          //customer push
          body.push({
            cell: { a: startRowCustomer+newIndex, b: 1, c: 0, d: 0, e: false },
            title: data.customerName,
            style: 1
          });

          //transaction total push
          body.push({
            cell: { a: startRowCustomer+newIndex, b: 2, c: 0, d: 0, e: false },
            title: data.transcationTotal,
            style: 1
          });

          //success push
          body.push({
            cell: { a: startRowCustomer+newIndex, b: 3, c: 0, d: 0, e: false },
            title: data.success,
            style: 1
          });

          //failed push
          body.push({
            cell: { a: startRowCustomer+newIndex, b: 4, c: 0, d: 0, e: false },
            title: data.failed,
            style: 1
          });
      })

      const lastRow = customerSampleData.length+1;

      const grandTotal = [
        {
           cell: { a: startRowCustomer+lastRow, b: 1, c: 0, d: 0, e: false },
           title: 'Grand Total',
           style: 0
         },
         {
           cell: { a: startRowCustomer+lastRow, b: 2, c: 0, d: 0, e: false },
           title: allTransaction,
           style: 0
         },
         {
           cell: { a: startRowCustomer+lastRow, b: 3, c: 0, d: 0, e: false },
           title: totalSuccess,
           style: 0
         },
         {
           cell: { a: startRowCustomer+lastRow, b: 4, c: 0, d: 0, e: false },
           title: totalFailed,
           style: 0
         }
     ];

     Array.prototype.push.apply(title, grandTotal);

     //end customer proccess

     //start operator process

     console.log(startRowCustomer+lastRow)

     let titleOperator = [
        {
          cell: { a: startRowCustomer+lastRow+2, b: 3, c: 0, d: 4, e: true },
          title: 'Status',
          style: 0
        },
        {
            cell: { a: startRowCustomer+lastRow+2, b: 1, c: startRowCustomer+lastRow+4, d: 0, e: true },
            title: 'Operator',
            style: 0
        },
        {
            cell: { a: startRowCustomer+lastRow+2, b: 2, c: startRowCustomer+lastRow+4, d: 0, e: true },
            title: 'Transaction Total',
            style: 0
        },
        {
            cell: { a: startRowCustomer+lastRow+3, b: 3, c: startRowCustomer+lastRow+3+1, d: 3, e: true },
            title: 'Success',
            style: 0
        },
        {
            cell: { a: startRowCustomer+lastRow+3, b: 4, c: startRowCustomer+lastRow+3+1, d: 4, e: true },
            title: 'Failed',
            style: 0
        }
      ];

      let bodyOperator = []

      Array.prototype.push.apply(title, titleOperator)
      

      const operatorSampleData = [
        {
            operatorName: 'Jasa Marga',
            transcationTotal: 1,
            success: 1,
            failed: 0
        },
        {
          operatorName: 'Jasa Marga 1',
          transcationTotal: 10,
          success: 5,
          failed: 5
      },
      {
          operatorName: 'Jasa Marga 2',
          transcationTotal: 20,
          success: 10,
          failed: 10
      },
      {
          operatorName: 'Makira 2',
          transcationTotal: 30,
          success: 10,
          failed: 10
      }
    ]

      let allTransactionOperator = 0;
      let totalSuccessOperator = 0;
      let totalFailedOperator = 0;
      const startRowOperator = (startRowCustomer*2)+lastRow;

      operatorSampleData.map((data, i) => {
        const newIndex = i+1;

        allTransactionOperator+=data.transcationTotal;
        totalSuccessOperator+=data.success;
        totalFailedOperator+=data.failed;

        //customer push
        bodyOperator.push({
          cell: { a: startRowOperator+newIndex, b: 1, c: 0, d: 0, e: false },
          title: data.operatorName,
          style: 1
        });

        //transaction total push
        bodyOperator.push({
          cell: { a: startRowOperator+newIndex, b: 2, c: 0, d: 0, e: false },
          title: data.transcationTotal,
          style: 1
        });

        //success push
        bodyOperator.push({
          cell: { a: startRowOperator+newIndex, b: 3, c: 0, d: 0, e: false },
          title: data.success,
          style: 1
        });

        //failed push
        bodyOperator.push({
          cell: { a: startRowOperator+newIndex, b: 4, c: 0, d: 0, e: false },
          title: data.failed,
          style: 1
        });
    });



    const lastRowOperator = operatorSampleData.length+1;

      const grandTotalOperator = [
        {
           cell: { a: startRowOperator+lastRowOperator, b: 1, c: 0, d: 0, e: false },
           title: 'Grand Total',
           style: 0
         },
         {
           cell: { a: startRowOperator+lastRowOperator, b: 2, c: 0, d: 0, e: false },
           title: allTransactionOperator,
           style: 0
         },
         {
           cell: { a: startRowOperator+lastRowOperator, b: 3, c: 0, d: 0, e: false },
           title: totalSuccessOperator,
           style: 0
         },
         {
           cell: { a: startRowOperator+lastRowOperator, b: 4, c: 0, d: 0, e: false },
           title: totalFailedOperator,
           style: 0
         }
     ];

     Array.prototype.push.apply(title, grandTotalOperator);
     Array.prototype.push.apply(body, bodyOperator);

     //end operator process


      body.push({
        cell: { a: 1, b: 1, c: 0, d: 0, e: false },
        title: `${formatStartDate}- ${formatEndDate}`,
        style: 1
      });

      await createExcel(title, body, style, 'Daily')

})();