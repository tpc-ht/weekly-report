import { FormLayout } from "@/components";
import { Button, Col, Form, Input, Row, Space, Typography, message } from "antd";
import { createStyles } from "antd-style";
import dayjs from "dayjs";
import ExcelJS from "exceljs";
import { useEffect, useRef, useState } from "react";
import CurrentWeekTable from "./CurrentWeekTable";
import NextWeekTable from "./NextWeekTable";
import { columns, excelColumnId, isDate, reportParse } from "./utils";
type dataType = {
  [key: string]: {
    project: string;
    module: string;
    content: string;
    startDate: string;
    endDate: string;
  };
};
const useStyles = createStyles(({ css }) => {
  return {
    main: css`
      padding: 0 20px 20px;
    `,
  };
});

export default () => {
  const [form] = Form.useForm();
  const { styles } = useStyles();
  const dateRef = useRef("");
  const [weekList, setWeekList] = useState<any[]>([]);
  const [weekTitle, setWeekTitle] = useState<string>("");
  const [nextWeekTitle, setNextWeekTitle] = useState<string>("");
  const [nextWeekList, setNextWeekList] = useState<any[]>([]);

  useEffect(() => {
    try {
      const value = localStorage.getItem("weekReport");
      value && form.setFieldsValue(JSON.parse(value));
    } catch (error) {}
  }, []);

  const currentWeekParse = (values: any) => {
    try {
      const strArr = values.currentWeekReport.split("\n");
      let project = values.project;
      let data: dataType = {};
      const dateCompare = (oldDate: string, newDate: string) => (dayjs(newDate).unix() > dayjs(oldDate).unix() ? newDate : oldDate);
      let currentDate = "";
      let msg = "";
      strArr.forEach((item: string, index: number) => {
        // 是否为时间分割
        let strItem = item.trim();
        if (strItem && !isDate(strItem)) {
          const obj = reportParse(strItem, project);
          if (data[obj.key]) {
            data[obj.key].project = obj.project;
            data[obj.key].module = obj.module;
            data[obj.key].content += obj.context + "\n";
            data[obj.key].endDate = dateCompare(data[obj.key].endDate, currentDate);
          } else {
            data[obj.key] = {
              project: obj.project,
              module: obj.module,
              content: obj.context + "\n",
              startDate: currentDate,
              endDate: currentDate,
            };
          }
        } else {
          currentDate = dayjs(strItem).format("YYYY-MM-DD");
          if (index === 0) {
            const sDate = dayjs(currentDate).format("YYYY.MM.DD");
            dateRef.current += sDate;
            msg += `${dayjs(strItem).format("YYYY年MM月")}工作周报：${sDate}`;
          }
        }
      });

      const eDate = dayjs(currentDate).format("YYYY.MM.DD");
      dateRef.current += `-${eDate}`;
      msg += `-${eDate}      `;
      msg += `部门：${values.department}      制表人：${values.leading}`;
      setWeekTitle(msg);
      const keys = Object.keys(data);

      let ls = [];
      for (let index = 0; index < keys.length; index++) {
        const key = keys[index];
        const item = data[key];
        ls.push({
          key: key,
          project: item.project,
          planType: "计划内",
          taskType: "需求",
          module: item.module,
          workContent: item.content,
          priority: "高",
          difficulty: "B",
          progress: "100%",
          state: "开发完成，提测",
          startDate: item.startDate,
          endDate: item.endDate,
          nextLeading: "测试",
          reason: "",
          leading: values.leading,
        });
      }
      setWeekList(ls);
    } catch (error) {
      message.error("请输入正确的本周日报格式");
    }
  };

  const nextWeekParse = (values: any) => {
    try {
      const strArr = values.nextWeekReport.split("\n");
      let project = values.project;
      let data: dataType = {};
      let currentDate = "";
      let endDate = "";

      strArr.forEach((item: string) => {
        // 是否为时间分割
        let strItem = item.trim();

        if (strItem) {
          if (!currentDate) {
            currentDate = strItem;
            endDate = dayjs(strItem.split("-")[1]).format("YYYY-MM-DD");
          } else {
            const obj = reportParse(strItem, project);
            if (data[obj.key]) {
              data[obj.key].project = obj.project;
              data[obj.key].module = obj.module;
              data[obj.key].content += obj.context + "\n";
            } else {
              data[obj.key] = {
                project: obj.project,
                module: obj.module,
                content: obj.context + "\n",
                startDate: "",
                endDate: endDate,
              };
            }
          }
        }
      });
      const title = ` 下周工作计划：${currentDate}`;
      setNextWeekTitle(title);
      const keys = Object.keys(data);

      let ls = [];
      for (let index = 0; index < keys.length; index++) {
        const key = keys[index];
        const item = data[key];
        ls.push({
          key,
          project: item.project,
          taskType: "需求",
          module: item.module,
          workContent: item.content,
          priority: "高",
          difficulty: "B",
          progress: "100%",
          state: "开发完成，提测",
          completionDate: item.endDate,
          note: "",
          leading: values.leading,
        });
      }
      setNextWeekList(ls);
      console.log("ls-next", ls);
    } catch (error) {
      console.log("error", error);
      message.error("请输入正确的下周计划格式");
    }
  };
  const onFinish = (values: any) => {
    localStorage.setItem("weekReport", JSON.stringify(values));
    currentWeekParse(values);
    nextWeekParse(values);
  };
  const setExcelGlobalStyle = (worksheet: ExcelJS.Worksheet, secondTableColNumber: number) => {
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
    const leftRegex = /B\d+/;
    const rightRegex01 = /P\d+/;

    const rightRegex02 = /N\d+/;
    worksheet.eachRow((row, rowNumber) => {
      row.eachCell((cell) => {
        if (excelColumnId.includes(cell.address[0])) {
          cell.style.font = { ...globalStyle.font, ...cell.style.font };
          cell.style.alignment = { ...globalStyle.alignment, ...cell.style.alignment };
          let border = { ...globalStyle.border, ...cell.style.border };
          if (leftRegex.test(cell.address)) {
            border = {
              ...border,
              left: { style: "medium" },
            };
          }
          if (rightRegex01.test(cell.address)) {
            border = {
              ...border,
              right: { style: "medium" },
            };
          }
          if (rightRegex02.test(cell.address) && rowNumber >= secondTableColNumber) {
            border = {
              ...border,
              right: { style: "medium" },
            };
          }
          cell.style.border = border;
        }
      });
    });
  };
  const setCurrentWeekTitle = (worksheet: ExcelJS.Worksheet) => {
    // 设置表格标题
    worksheet.mergeCells("B2:P2");
    const titleCell = worksheet.getCell("B2:P2");
    titleCell.value = "研发部工作周报";
    titleCell.font = {
      size: 18,
      bold: true,
    };
    titleCell.alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    titleCell.border = {
      top: { style: "medium" },
    };
    worksheet.mergeCells("B3:P3");
    const titleCell2 = worksheet.getCell("B3:P3");
    titleCell2.value = weekTitle;
    titleCell2.font = {
      bold: true,
    };
    titleCell2.alignment = {
      vertical: "middle",
      horizontal: "left",
    };
  };
  const setCurrentWeekBody = (worksheet: ExcelJS.Worksheet) => {
    worksheet.columns = columns;
    worksheet.addRow(["", "序号", "项目名称", "计划类型", "任务类型", "功能模块", "工作内容", "优先级", "难度", "完成进度", "任务状态", "计划开始日期", "计划结束日期", "下阶段负责人", "未完成原因分析", "负责人"], "A10");
    for (let index = 0; index < weekList.length; index++) {
      const { key, project, planType, taskType, module, workContent, priority, difficulty, progress, state, startDate, endDate, nextLeading, reason, leading } = weekList[index];
      worksheet.addRow(["", index + 1, project, planType, taskType, module, workContent.trim(), priority, difficulty, progress, state, startDate, endDate, nextLeading, reason, leading]);
    }
  };
  const setCurrentWeekFooter = (worksheet: ExcelJS.Worksheet, cIndex: number) => {
    // 表底部
    const footerAddress01 = `B${cIndex}:C${cIndex}`;
    const footerValueAddress01 = `D${cIndex}:P${cIndex}`;
    worksheet.mergeCells(footerAddress01);
    const footerCell01 = worksheet.getCell(footerAddress01);
    footerCell01.value = "工作总结（必填）：";
    worksheet.mergeCells(footerValueAddress01);
    const footerValueCell01 = worksheet.getCell(footerValueAddress01);
    footerValueCell01.value = "按计划进行";
    footerValueCell01.style.alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    let cIndex02 = cIndex + 1;
    const footerAddress02 = `B${cIndex02}:C${cIndex02}`;
    const footerValueAddress02 = `D${cIndex02}:P${cIndex02}`;
    worksheet.mergeCells(footerAddress02);
    const footerCell02 = worksheet.getCell(footerAddress02);
    footerCell02.value = "风险及建议";
    footerCell02.style.border = {
      bottom: { style: "medium" },
    };
    worksheet.mergeCells(footerValueAddress02);
    const footerValueCell02 = worksheet.getCell(footerValueAddress02);
    footerValueCell02.style.alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    footerValueCell02.style.border = {
      bottom: { style: "medium" },
    };
  };

  const setNextWeekHeader = (worksheet: ExcelJS.Worksheet, cIndex: number) => {
    worksheet.mergeCells(`B${cIndex}:N${cIndex}`);
    const titleCell2 = worksheet.getCell(`B${cIndex}:N${cIndex}`);
    titleCell2.value = ` 下周工作计划：24.07.08-24.07.12`;
    titleCell2.font = {
      bold: true,
    };
    titleCell2.alignment = {
      vertical: "middle",
      horizontal: "left",
    };
    titleCell2.border = {
      top: { style: "medium" },
    };
  };
  const setNextWeekBody = (worksheet: ExcelJS.Worksheet, cIndex: number) => {
    worksheet.addRow(["", "序号", "项目名称", "任务类型", "功能模块", "功能模块", "工作内容", "优先级", "难度", "计划完成进度", "计划测试状态", "计划完成日期", "补充说明", "责任人"]);

    worksheet.mergeCells(`E${cIndex}:F${cIndex}`);
    let cIndex01 = cIndex;
    for (let index = 0; index < nextWeekList.length; index++) {
      const { project, taskType, module, workContent, priority, difficulty, progress, state, completionDate, note, leading } = nextWeekList[index];
      cIndex01 += 1;
      worksheet.addRow(["", index + 1, project, taskType, module, module, workContent.trim(), priority, difficulty, progress, state, completionDate, note, leading]);
      worksheet.mergeCells(`E${cIndex01}:F${cIndex01}`);
    }
  };
  const setNextWeekFooter = (worksheet: ExcelJS.Worksheet, cIndex: number) => {
    const footerAddress03 = `B${cIndex}:C${cIndex}`;
    const footerValueAddress03 = `D${cIndex}:N${cIndex}`;
    worksheet.mergeCells(footerAddress03);
    const footerCell03 = worksheet.getCell(footerAddress03);
    footerCell03.value = "风险问题";
    worksheet.mergeCells(footerValueAddress03);

    let cIndex01 = cIndex + 1;
    const footerAddress04 = `B${cIndex01}:C${cIndex01}`;
    const footerValueAddress04 = `D${cIndex01}:N${cIndex01}`;
    worksheet.mergeCells(footerAddress04);
    const footerCell04 = worksheet.getCell(footerAddress04);
    footerCell04.value = "沟通协调";
    footerCell04.style.border = {
      bottom: { style: "medium" },
    };
    worksheet.mergeCells(footerValueAddress04);
    const footerValueCell04 = worksheet.getCell(footerValueAddress04);
    footerValueCell04.style.alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    footerValueCell04.style.border = {
      bottom: { style: "medium" },
    };
  };
  const exportFile = () => {
    const value = form.getFieldsValue();
    // 创建工作簿
    const workbook = new ExcelJS.Workbook();
    // 添加工作表
    const worksheet = workbook.addWorksheet("sheet1");
    worksheet.addRow([]);
    setCurrentWeekTitle(worksheet);
    setCurrentWeekBody(worksheet);
    const cIndex01 = 5 + weekList.length;
    setCurrentWeekFooter(worksheet, cIndex01);

    const cIndex02 = cIndex01 + 3;
    setNextWeekHeader(worksheet, cIndex02);

    let cIndex03 = cIndex02 + 1;
    setNextWeekBody(worksheet, cIndex03);

    let cIndex04 = cIndex03 + nextWeekList.length + 1;
    setNextWeekFooter(worksheet, cIndex04);

    // 全局样式
    setExcelGlobalStyle(worksheet, cIndex02);
    // 导出表格
    workbook.xlsx.writeBuffer().then((buffer) => {
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = `${value.department ? value.department + "_" : ""}${value.leading ? value.leading + "_" : ""}${dateRef.current ? dateRef.current + "_" : ""}工作周报.xlsx`;
      link.click();
      URL.revokeObjectURL(link.href); // 下载完成释放掉blob对象
    });
  };
  return (
    <div className={styles.main}>
      <Typography.Title level={3} style={{ textAlign: "center" }}>
        周报生成工具
      </Typography.Title>
      <FormLayout>
        <Form
          onFinish={onFinish}
          form={form}
          autoComplete="off"
          initialValues={{
            project: "crm",
            department: "前端部",
            summary: "计划进行",
          }}
        >
          <Row gutter={16}>
            <Col span={8}>
              <Form.Item
                label="默认项目名称"
                rules={[
                  {
                    required: true,
                  },
                ]}
                name={"project"}
              >
                <Input placeholder="请输入" />
              </Form.Item>
            </Col>
            <Col span={8}>
              <Form.Item
                label="部门"
                rules={[
                  {
                    required: true,
                  },
                ]}
                name={"department"}
              >
                <Input placeholder="请输入" />
              </Form.Item>
            </Col>
            <Col span={8}>
              <Form.Item
                label="负责人"
                rules={[
                  {
                    required: true,
                  },
                ]}
                name={"leading"}
              >
                <Input placeholder="请输入" />
              </Form.Item>
            </Col>
          </Row>

          <Form.Item
            label="工作总结"
            rules={[
              {
                required: true,
              },
            ]}
            name={"summary"}
          >
            <Input placeholder="请输入" />
          </Form.Item>
          <Form.Item
            label="本周日报"
            rules={[
              {
                required: true,
              },
            ]}
            name={"currentWeekReport"}
            tooltip={
              <div>
                格式：
                <div>2024-01-01</div>
                <div>1. [项目名称] 模块名称 工作内容</div>
                <div>2024-01-02</div>
                <div>1. [项目名称] 模块名称 工作内容</div>
              </div>
            }
          >
            <Input.TextArea rows={6} placeholder="请输入" />
          </Form.Item>
          <Form.Item
            label="下周计划"
            rules={[
              {
                required: true,
              },
            ]}
            name={"nextWeekReport"}
            tooltip={
              <div>
                格式：
                <div>2024.07.08-2024.07.12</div>
                <div>1. [项目名称] 模块名称 工作内容</div>
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
            <Button disabled={!(weekList.length && nextWeekList.length)} style={{ margin: "0 auto" }} onClick={exportFile}>
              导出
            </Button>
          </Space>
        </Form>
      </FormLayout>
      <Typography.Title level={4}>本周周报预览</Typography.Title>
      <Typography.Paragraph>{weekTitle}</Typography.Paragraph>
      <CurrentWeekTable data={weekList} />
      <Typography.Title level={4}>下周报预览</Typography.Title>
      <Typography.Paragraph>{nextWeekTitle}</Typography.Paragraph>
      <NextWeekTable data={nextWeekList} />
    </div>
  );
};
