import type { TableProps } from "antd";
import { Button, Space, Table, Tag } from "antd";
import xlsx from "xlsx";
import styles from "./index.less";

interface DataType {
  key: string;
  name: string;
  age: number;
  address: string;
  tags: string[];
}


const columns: TableProps<DataType>["columns"] = [
  {
    title: "Name",
    dataIndex: "name",
    key: "name",
    render: (text) => <a>{text}</a>,
  },
  {
    title: "Age",
    dataIndex: "age",
    key: "age",
  },
  {
    title: "Address",
    dataIndex: "address",
    key: "address",
  },
  {
    title: "Tags",
    key: "tags",
    dataIndex: "tags",
    render: (_, { tags }) => (
      <>
        {tags.map((tag) => {
          let color = tag.length > 5 ? "geekblue" : "green";
          if (tag === "loser") {
            color = "volcano";
          }
          return (
            <Tag color={color} key={tag}>
              {tag.toUpperCase()}
            </Tag>
          );
        })}
      </>
    ),
  },
  {
    title: "Action",
    key: "action",
    render: (_, record) => (
      <Space size="middle">
        <a>Invite {record.name}</a>
        <a>Delete</a>
      </Space>
    ),
  },
];

const data: DataType[] = [
  {
    key: "1",
    name: "John Brown",
    age: 32,
    address: "New York No. 1 Lake Park",
    tags: ["nice", "developer"],
  },
  {
    key: "2",
    name: "Jim Green",
    age: 42,
    address: "London No. 1 Lake Park",
    tags: ["loser"],
  },
  {
    key: "3",
    name: "Joe Black",
    age: 32,
    address: "Sydney No. 1 Lake Park",
    tags: ["cool", "teacher"],
  },
];
 // 将一个sheet转成最终的excel文件的blob对象，然后利用URL.createObjectURL下载
 function sheet2blob(sheet:any, sheetName='sheet1') {
    var workbook:any = {
        SheetNames: [sheetName],
        Sheets: {}
    };
    workbook.Sheets[sheetName] = sheet;
    var wbout = xlsx.write(workbook, {
        bookType: 'xlsx', // 要生成的文件类型
        bookSST: false, // 是否生成Shared String Table，官方解释是，如果开启生成速度会下降，但在低版本IOS设备上有更好的兼容性
        type: 'binary'
    });
    var blob = new Blob([s2ab(wbout)], {type:"application/octet-stream"});
    // 字符串转ArrayBuffer
    function s2ab(s:any) {
        var buf = new ArrayBuffer(s.length);
        var view = new Uint8Array(buf);
        for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
    }
    return blob;
}
const exportTable = () => {
  //
  //
  //
  //
  //
  //
 console.log("xlsx",xlsx);
//  if(1)return
  
//   var blob = sheet2blob(xlsx.utils.table_to_sheet(document.querySelector("#table_report")))
//   var link = window.URL.createObjectURL(blob); 
//   var a = document.createElement("a");    //创建a标签
//   a.download = "企业反映问题诉求汇总表.xlsx";                //设置被下载的超链接目标（文件名）
//   a.href = link;                            //设置a标签的链接
//   document.body.appendChild(a);            //a标签添加到页面
//   a.click();                                //设置a标签触发单击事件
//   document.body.removeChild(a);            //移除a标签
/*  */
  //   const tableDom = document.querySelector("#table_container table")?.outerHTML;
  const tableDom = document.querySelector("#table_report")?.outerHTML;
  
  if (!tableDom) return;

  var blob = new Blob([tableDom], { type: "text/plain;charset=utf-8" }); //解决中文乱码问题
  blob = new Blob([String.fromCharCode(0xfeff), blob], { type: blob.type });
  //设置链接
  var link = window.URL.createObjectURL(blob);
  var a = document.createElement("a"); //创建a标签
  a.download = "企业反映问题诉求汇总表.xls"; //设置被下载的超链接目标（文件名）
  a.href = link; //设置a标签的链接
  document.body.appendChild(a); //a标签添加到页面
  a.click(); //设置a标签触发单击事件
  document.body.removeChild(a); //移除a标签
};

export default () => (
  <div id="table_container">
    <Button onClick={exportTable}>表格导出</Button>

    <table id="table_report" className={styles.table}  style={{ fontFamily: "微软雅黑",fontSize: 14 }} border={1}>
      <caption style={{ textAlign: "center" }}>
        <h3>研发部工作周报</h3>
      </caption>
        <colgroup>
            <col width={50}/>
        </colgroup>
      <thead >
      <tr>
        <td colSpan={15} >
            <div style={{display:"flex",alignItems:"center",gap:100}}>
                <div>2024年06月工作周报：24.06.24-24.06.28</div>
                <div>部门：前端组</div>
                <div>制表人：唐鹏程</div>
            </div>
        </td>
      </tr>
      </thead>
      <tbody > 
      <tr>
        <td align="center">序号</td>
        <td >合计</td>
        <td >合计</td>
        <td width={100}>companyNum</td>
        <td width={100}>companyNum</td>
        <td width={100}>questionNum</td>
        <td width={100}>type0Num</td>
        <td width={100}>type1Num</td>
        <td width={100}>type2Num</td>
        <td width={100}>type3Num</td>
        <td width={100}>type4Num</td>
        <td width={100}>type5Num</td>
        <td width={100}>type6Num</td>
        <td width={100}>type7Num</td>
        <td width={100}>type8Num</td>
        <td width={100}></td>
      </tr>
      {/* <tr>
        <td width={50} style={{minWidth:'50px'}}>1</td>
        <td width={100}>市直</td>
        <td width={100}>shizhiList.companyNum</td>
        <td width={100}>shizhiList.companyNum</td>
        <td width={100}>shizhiList.questionNum</td>
        <td width={100}>shizhiList.type0Num</td>
        <td width={100}>shizhiList.type1Num</td>
        <td width={100}>shizhiList.type2Num</td>
        <td width={100}>shizhiList.type3Num</td>
        <td width={100}>shizhiList.type4Num</td>
        <td width={100}>shizhiList.type5Num</td>
        <td width={100}>shizhiList.type6Num</td>
        <td width={100}>shizhiList.type7Num</td>
        <td width={100}>shizhiList.type8Num</td>
        <td width={100}></td>
      </tr> */}


      </tbody>
    </table>
    <Table bordered columns={columns} dataSource={data} />
  </div>
);
