import { Button, Form, Input, Space, Typography, message } from "antd";
import dayjs from "dayjs";
import ExcelJS from "exceljs";
import { useState } from "react";
type dataType = {
  [key: string]: {
    content: string;
    startDate: string;
    endDate: string;
  };
};
export default () => {
  const [form] = Form.useForm();
  const [content, setContent] = useState("");
  const [data, setData] = useState<dataType>({});

  const strFormat = (str: string, defaultProject: string) => {
    console.log("str", str);

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
  const exportFile = () => {
    // 创建工作簿
    const workbook = new ExcelJS.Workbook();
    // 添加工作表
    const worksheet = workbook.addWorksheet("sheet1");

    // 设置表头
    worksheet.columns = [
      {
        header: "序号",
        key: "index",
        width: 5,
        style: {
          alignment: {
            horizontal: "center",
          },
        },
      },
      {
        header: "项目名称",
        key: "orgName",
        width: 10,
        style: {
          alignment: {
            horizontal: "center",
          },
        },
      },
      {
        header: "计划类型",
        key: "planType",
        width: 10,
        style: {
          alignment: {
            horizontal: "center",
          },
        },
      },
      {
        header: "任务类型",
        key: "taskType",
        width: 10,
        style: {
          alignment: {
            horizontal: "center",
          },
        },
      },
      {
        header: "功能模块",
        key: "module",
        width: 15,
        style: {
          alignment: {
            horizontal: "center",
          },
        },
      },
      {
        header: "工作内容",
        key: "content",
        width: 50,
      },

      {
        header: "优先级",
        key: "priority",
        width: 10,
        style: {
          alignment: {
            horizontal: "center",
          },
        },
      },
      {
        header: "难度",
        key: "difficulty",
        width: 10,
        style: {
          alignment: {
            horizontal: "center",
          },
        },
      },
      {
        header: "完成进度",
        key: "progress",
        width: 10,
        style: {
          alignment: {
            horizontal: "center",
          },
        },
      },
      {
        header: "任务状态",
        key: "state",
        width: 15,
        style: {
          alignment: {
            horizontal: "center",
          },
        },
      },
      {
        header: "计划开始日期",
        key: "startDate",
        width: 12,
        style: {
          alignment: {
            horizontal: "center",
          },
        },
      },
      {
        header: "计划结束日期",
        key: "endDate",
        width: 12,
        style: {
          alignment: {
            horizontal: "center",
          },
        },
      },
      {
        header: "下阶段负责人",
        key: "nextLeader",
        width: 10,
        style: {
          alignment: {
            horizontal: "center",
          },
        },
      },
      {
        header: "未完成原因分析",
        key: "note",
        width: 15,
        style: {
          alignment: {
            horizontal: "center",
          },
        },
      },
      {
        header: "负责人",
        key: "leader",
        width: 10,
        style: {
          alignment: {
            horizontal: "center",
          },
        },
      },
    ];

    // 添加表体数据
    const keys = Object.keys(data);
    let orgName = form.getFieldValue("orgName");
    for (let index = 0; index < keys.length; index++) {
      const key = keys[index];
      rowStyleFormat(worksheet, index + 1);
      worksheet.addRow({
        index: index + 1,
        orgName: orgName,
        planType: "计划内",
        taskType: "需求",
        module: key,
        content: data[key].content.trim(),
        priority: "高",
        difficulty: "B",
        progress: "100%",
        state: "开发完成，提测",
        startDate: data[key].startDate.trim(),
        endDate: data[key].endDate.trim(),
        nextLeader: "测试",
        note: "",
        leader: "唐***",
      });
    }

    rowStyleFormat(worksheet, keys.length + 1);
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
      </Space>
      <Form.Item label="结果" noStyle shouldUpdate style={{ minHeight: 300 }}>
        {() => (
          <Typography>
            <pre style={{ minHeight: 300 }}>{content}</pre>
          </Typography>
        )}
      </Form.Item>
    </Form>
  );
};
