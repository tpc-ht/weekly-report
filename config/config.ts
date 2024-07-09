import { defineConfig } from "umi";

export default defineConfig({

    routes: [
        { path: "/", component: "index" },
    ],
    npmClient: 'pnpm',
    title: '周报生成工具',
    history: {
        type: 'hash',
    },
    publicPath: process.env.NODE_ENV === 'production' ? './' : '/'

});






