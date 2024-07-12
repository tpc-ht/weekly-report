import type { TableProps } from "antd";
import { Table } from "antd";

interface DataType {
  key: string;
  name: string;
  age: number;
  address: string;
  tags: string[];
}

const columns: TableProps<DataType>["columns"] = [
  {
    title: "序号",
    dataIndex: "key",
    rowScope: "row",
    width: 60,
    render: (text, item, index) => index + 1,
  },
  {
    dataIndex: "project",
    title: "项目名称",
    width: 80,
  },
  {
    title: "计划类型",
    width: 80,
    dataIndex: "planType",
  },
  {
    title: "任务类型",
    width: 80,
    dataIndex: "taskType",
  },
  {
    dataIndex: "module",
    width: 120,
    title: "功能模块",
  },
  {
    dataIndex: "workContent",
    width: 300,
    title: "工作内容",
    render: (text: string) => {
      return (
        <div>
          {text.split("\n").map((e) => (
            <div>{e}</div>
          ))}
        </div>
      );
    },
  },
  {
    dataIndex: "priority",
    width: 60,
    title: "优先级",
  },
  {
    dataIndex: "difficulty",
    width: 60,
    title: "难度",
  },
  {
    dataIndex: "progress",
    width: 60,
    title: "进度",
  },
  {
    dataIndex: "state",
    width: 120,
    title: "任务状态",
  },
  {
    dataIndex: "endDate",
    width: 120,
    title: "计划完成日期",
  },
  {
    dataIndex: "endDate",
    width: 120,
    title: "实际完成日期",
  },
  {
    dataIndex: "nextLeading",
    width: 120,
    title: "下阶段负责人",
  },
  {
    dataIndex: "reason",
    width: 120,
    title: "未完成原因分析",
  },
  {
    dataIndex: "leading",
    width: 80,
    title: "负责人",
  },
];

type TablePropsType = {
  data: any;
};

export default ({ data = [] }: TablePropsType) => <Table rowKey="key" size="small" bordered scroll={{ x: 1600 }} columns={columns} dataSource={data} />;
