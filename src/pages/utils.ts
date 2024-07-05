import { Column } from "exceljs";

export const columns: Partial<Column>[] = [
    {
        header: "",
        key: "",
        // mergeCells
        width: 2,
    },
    {
        // header: "序号",
        key: "index",
        width: 5,
        style: {
            alignment: {
                horizontal: "center",
            },
        },
    },
    {
        // header: "项目名称",
        key: "orgName",
        width: 10,
        style: {
            alignment: {
                horizontal: "center",
            },
        },
    },
    {
        // header: "计划类型",
        key: "planType",
        width: 10,
        style: {
            alignment: {
                horizontal: "center",
            },
        },
    },
    {
        // header: "任务类型",
        key: "taskType",
        width: 10,
        style: {
            alignment: {
                horizontal: "center",
            },
        },
    },
    {
        // header: "功能模块",
        key: "module",
        width: 15,
        style: {
            alignment: {
                horizontal: "center",
            },
        },
    },
    {
        // header: "工作内容",
        key: "content",
        width: 50,
    },

    {
        // header: "优先级",
        key: "priority",
        width: 10,
        style: {
            alignment: {
                horizontal: "center",
            },
        },
    },
    {
        // header: "难度",
        key: "difficulty",
        width: 10,
        style: {
            alignment: {
                horizontal: "center",
            },
        },
    },
    {
        // header: "完成进度",
        key: "progress",
        width: 10,
        style: {
            alignment: {
                horizontal: "center",
            },
        },
    },
    {
        // header: "任务状态",
        key: "state",
        width: 15,
        style: {
            alignment: {
                horizontal: "center",
            },
        },
    },
    {
        // header: "计划开始日期",
        key: "startDate",
        width: 12,
        style: {
            alignment: {
                horizontal: "center",
            },
        },
    },
    {
        // header: "计划结束日期",
        key: "endDate",
        width: 12,
        style: {
            alignment: {
                horizontal: "center",
            },
        },
    },
    {
        // header: "下阶段负责人",
        key: "nextLeader",
        width: 12,
        style: {
            alignment: {
                horizontal: "center",
            },
        },
    },
    {
        // header: "未完成原因分析",
        key: "note",
        width: 15,
        style: {
            alignment: {
                horizontal: "center",
            },
        },
    },
    {
        // header: "负责人",
        key: "leader",
        width: 10,
        style: {
            alignment: {
                horizontal: "center",
            },
        },
    },
]