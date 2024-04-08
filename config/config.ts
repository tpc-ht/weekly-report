import { defineConfig } from "umi";

export default defineConfig({

    routes: [
        { path: "/", component: "index" },
    ],
    npmClient: 'pnpm',
    title: '日报解析',
    history: {
        type: 'hash',
    },
    publicPath: process.env.NODE_ENV === 'production' ? './' : '/'

});






