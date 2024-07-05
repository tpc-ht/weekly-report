import { Button, Form, Input, Space, Typography, message } from "antd";
import dayjs from "dayjs";
import ExcelJS from "exceljs";
import { useState } from "react";
import Table from "./Table";
import { columns } from "./utils";
type dataType = {
  [key: string]: {
    content: string;
    startDate: string;
    endDate: string;
  };
};
export default () => {
  const [form] = Form.useForm();
  const [content, setContent] = useState(`投票管理 2024-06-24 - 2024-06-27
投票设置交互调试
作品设置表单功能开发与交互调试
投票接口开发与接口联调
预览模块样式优化与表单交互调整
主表操作项功能开发联调
分组显示问题调试

jpos 2024-06-26 - 2024-06-27
新用户登录密码重置开发
系统参数配置接口调试，新增密码重置验证期限设置

投票活动 2024-06-26 - 2024-06-26
新增表单交互值丢失问题修复

抽奖活动 2024-06-27 - 2024-06-27
表字段显示调试

系统 2024-06-28 - 2024-06-28
富文本组件内容验证调试
拦截器异常处理调试`);
  const [data, setData] = useState<dataType>({
    投票管理: { content: "投票设置交互调试\n作品设置表单功能开发与交互调试\n投票接口开发与接口联调\n预览模块样式优化与表单交互调整\n主表操作项功能开发联调\n分组显示问题调试\n", startDate: "2024-06-24", endDate: "2024-06-27" },
    jpos: { content: "新用户登录密码重置开发\n系统参数配置接口调试，新增密码重置验证期限设置\n", startDate: "2024-06-26", endDate: "2024-06-27" },
    投票活动: { content: "新增表单交互值丢失问题修复\n", startDate: "2024-06-26", endDate: "2024-06-26" },
    抽奖活动: { content: "表字段显示调试\n", startDate: "2024-06-27", endDate: "2024-06-27" },
    系统: { content: "富文本组件内容验证调试\n拦截器异常处理调试\n", startDate: "2024-06-28", endDate: "2024-06-28" },
  });

  const strFormat = (str: string, defaultProject: string) => {
    const strArr = str.split(".")[1].trim().split(" ");
    let data = {
      project: defaultProject,
      menu: "",
      context: "",
      isUse: false,
    };
    if (strArr.length === 2) {
      data.menu = strArr[0];
      data.context = strArr[1];
    } else if (strArr.length === 3) {
      data.project = strArr[0];
      data.menu = strArr[1];
      data.context = strArr[2];
    }
    return data;
  };

  const onFinish = (values: any) => {
    try {
      const strArr = values.content.split("\n");
      let project = "商管";
      let data: dataType = {};
      const isDate = (date: string) => {
        if (!data) return false;
        const regex = /^\d\d\d\d[-](\d{1,2})[-](\d{1,2})$/;
        return regex.test(date);
      };
      const dateCompare = (oldDate: string, newDate: string) => (dayjs(newDate).unix() > dayjs(oldDate).unix() ? newDate : oldDate);
      let currentDate = "";
      strArr.forEach((item: string) => {
        // 是否为时间分割
        let strItem = item.trim();

        if (strItem && !isDate(strItem)) {
          const obj = strFormat(strItem, project);
          if (obj.menu) {
            if (data[obj.menu]) {
              data[obj.menu].content += obj.context + "\n";
              data[obj.menu].endDate = dateCompare(data[obj.menu].endDate, currentDate);
            } else {
              data[obj.menu] = {
                content: obj.context + "\n",
                startDate: currentDate,
                endDate: currentDate,
              };
            }
          }
        } else {
          currentDate = strItem;
        }
      });

      const keys = Object.keys(data);
      let text = "";
      for (let index = 0; index < keys.length; index++) {
        const key = keys[index];
        text += `${key} ${data[key].startDate} - ${data[key].endDate}\n${data[key].content}\n`;
      }
      console.log("text.trim()", JSON.stringify(data));
      setData(data);
      setContent(text.trim());
    } catch (error) {
      message.error("请输入正确的格式");
    }
  };
  const rowStyleFormat = (worksheet: ExcelJS.Worksheet, index: number) => {
    let rows = worksheet.getRow(index);
    rows.font = {
      name: "微软雅黑",
      size: 10,
    };
    rows.alignment = {
      vertical: "middle",
    };
    rows.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };
  };
  const setExcelGlobalStyle = (worksheet: ExcelJS.Worksheet) => {
    // 创建一个样式对象，设置所需的样式属性
    const globalStyle: any = {
      font: {
        name: "微软雅黑",
        size: 10,
      },
      alignment: {
        vertical: "middle",
        wrapText: true,
      },
      border: {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      },
    };
    worksheet.eachRow((row, rowNumber) => {
      row.eachCell((cell, colNumber) => {
        console.log("cell", cell);
        // const num = cell.address.substring(1);
        // if (num !== "1" && num !== "2") {
        //   // cell.column = colNumber
        // }
        cell.style.font = { ...globalStyle.font, ...cell.style.font };
        cell.style.alignment = { ...globalStyle.alignment, ...cell.style.alignment };
        cell.style.border = { ...globalStyle.border, ...cell.style.border };
      });
    });
  };
  const setTitle = (worksheet: ExcelJS.Worksheet) => {
    // 设置表格标题
    worksheet.mergeCells("B1:P1");
    const titleCell = worksheet.getCell("B1:P1");
    titleCell.value = "研发部工作周报";
    titleCell.font = {
      size: 18,
      bold: true,
    };
    titleCell.alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    worksheet.mergeCells("B2:P2");
    const titleCell2 = worksheet.getCell("B2:P2");
    titleCell2.value = `2024年06月工作周报：24.06.24-24.06.28                                                                      部门：       前端组              制表人：      唐鹏程`;
    titleCell2.font = {
      bold: true,
    };
    titleCell2.alignment = {
      vertical: "middle",
      horizontal: "left",
    };
  };
  const exportFileTest = async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Sheet1");

    // 设置标题
    worksheet.getCell("A1").value = "姓名";
    worksheet.getCell("B1").value = "年龄";
    worksheet.getCell("C1").value = "性别";

    // 添加行数据
    // worksheet.addRow(["张三", 25, "男"]);
    // worksheet.addRow(["李四", 30, "女"]);
    worksheet.addRow({ id: 1, name: "John Doe", dob: new Date(1970, 1, 1) });
    worksheet.addRow({ id: 2, name: "Jane Doe", dob: new Date(1965, 1, 7) });

    // 保存 Excel 文件
    // await workbook.xlsx.writeFile("example.xlsx");

    // 导出表格
    workbook.xlsx.writeBuffer().then((buffer) => {
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = "周报1.xlsx";
      link.click();
      URL.revokeObjectURL(link.href); // 下载完成释放掉blob对象
    });
  };
  const exportFile = () => {
    // 创建工作簿
    const workbook = new ExcelJS.Workbook();
    // 添加工作表
    const worksheet = workbook.addWorksheet("sheet1", {
      headerFooter: { firstHeader: "Hello Exceljs", firstFooter: "Hello World" },
    });
    setTitle(worksheet);
    worksheet.columns = columns;
    // worksheet.addRow();
    const headerStyle = {
      font: { bold: true },
      alignment: { horizontal: "center" },
    };
    worksheet.addRow(["", "序号", "项目名称", "计划类型", "任务类型", "功能模块", "工作内容", "优先级", "难度", "完成进度", "任务状态", "计划开始日期", "计划结束日期", "下阶段负责人", "未完成原因分析", "负责人"], "A10");
    // 添加表体数据
    const keys = Object.keys(data);
    let orgName = form.getFieldValue("orgName");
    for (let index = 0; index < keys.length; index++) {
      const key = keys[index];
      // rowStyleFormat(worksheet, index + 1);
      worksheet.addRow(["", index + 1, orgName, "计划内", "需求", key, data[key].content.trim(), "高", "B", "100%", "开发完成，提测", data[key].startDate.trim(), data[key].endDate.trim(), "测试", "", "唐***"]);

      // {
      //   index: index + 1,
      //   orgName: orgName,
      //   planType: "计划内",
      //   taskType: "需求",
      //   module: key,
      //   content: data[key].content.trim(),
      //   priority: "高",
      //   difficulty: "B",
      //   progress: "100%",
      //   state: "开发完成，提测",
      //   startDate: data[key].startDate.trim(),
      //   endDate: data[key].endDate.trim(),
      //   nextLeader: "测试",
      //   note: "",
      //   leader: "唐***",
      // }
    }

    // rowStyleFormat(worksheet, keys.length + 1);
    setExcelGlobalStyle(worksheet);

    console.log("worksheet", worksheet);
    // const bp = worksheet.model.getRanges("B2:P8");
    // const bp = workbook.definedNames.getRanges("sheet1");
    // 设置 A1:B2 单元格的样式
    // worksheet.getRanges("B2:P8").font = {
    //   bold: true,
    //   color: { argb: "FF0000" },
    // };
    // console.log("bp", bp);

    // 导出表格
    workbook.xlsx.writeBuffer().then((buffer) => {
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = "周报.xlsx";
      link.click();
      URL.revokeObjectURL(link.href); // 下载完成释放掉blob对象
    });
  };
  return (
    <div>
      <Form
        onFinish={onFinish}
        form={form}
        layout="vertical"
        name="dynamic_form_complex"
        style={{
          maxWidth: 600,
          margin: "0 auto",
        }}
        autoComplete="off"
        initialValues={{
          items: [{}],
          orgName: "crm",
          content: `2024-06-24
        1. 投票管理 投票设置交互调试
        2. 投票管理 作品设置表单功能开发与交互调试
        2024-06-25
        1. 投票管理 投票接口开发与接口联调
        2. 投票管理 预览模块样式优化与表单交互调整
        3. 投票管理 主表操作项功能开发联调
        2024-06-26
        1. jpos 新用户登录密码重置开发
        2. 投票活动 新增表单交互值丢失问题修复
        2024-06-27
        1. jpos 系统参数配置接口调试，新增密码重置验证期限设置
        2. 抽奖活动 表字段显示调试
        3. 投票管理 分组显示问题调试
        2024-06-28
        1. 系统 富文本组件内容验证调试
        2. 系统 拦截器异常处理调试`,
        }}
      >
        <Form.Item label="项目名称" name={"orgName"}>
          <Input placeholder="请输入" />
        </Form.Item>
        <Form.Item
          label="日报内容"
          rules={[
            {
              required: true,
            },
          ]}
          name={"content"}
          tooltip={
            <div>
              格式：
              <div>2024-04-01</div>
              <div>1. 人员管理 功能开发</div>
              <div>2. 角色管理 XXXX开发</div>
            </div>
          }
        >
          <Input.TextArea rows={6} placeholder="请输入" />
        </Form.Item>
        <Space
          style={{
            width: "100%",
            justifyContent: "flex-end",
          }}
        >
          <Button type="primary" style={{ margin: "0 auto" }} htmlType="submit">
            解析
          </Button>
          <Button disabled={!Object.keys(data).length} style={{ margin: "0 auto" }} onClick={exportFile}>
            导出
          </Button>
          <Button disabled={!Object.keys(data).length} style={{ margin: "0 auto" }} onClick={exportFileTest}>
            导出2
          </Button>
        </Space>
        <Form.Item label="结果" noStyle shouldUpdate style={{ minHeight: 300 }}>
          {() => (
            <Typography>
              <pre style={{ minHeight: 300 }}>{content}</pre>
            </Typography>
          )}
        </Form.Item>
      </Form>
      <Table />
    </div>
  );
};
