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
    width: 60,
    render: (text, item, index) => index + 1,
  },
  {
    dataIndex: "project",
    title: "项目名称",
    width: 80,
  },
  {
    title: "任务类型",
    width: 80,
    dataIndex: "taskType",
  },
  {
    dataIndex: "module",
    width: 80,
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
    width: 120,
    title: "计划完成进度",
  },
  {
    dataIndex: "state",
    width: 120,
    title: "计划测试状态",
  },
  {
    dataIndex: "completionDate",
    width: 120,
    title: "计划完成日期",
  },
  {
    dataIndex: "note",
    width: 80,
    title: "补充说明",
  },
  {
    dataIndex: "leading",
    width: 80,
    title: "责任人",
  },
];

type TablePropsType = {
  data: any[];
};

export default ({ data = [] }: TablePropsType) => <Table rowKey="key" size="small" bordered scroll={{ x: 1300 }} columns={columns} dataSource={data} />;
