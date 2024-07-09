import { ConfigProvider } from "antd";
import locale from "antd/es/locale/zh_CN";
import dayjs from "dayjs";
import "dayjs/locale/zh-cn"; // 当时不加日期内部的年月没有汉化
import { Outlet } from "umi";
dayjs.locale("zh-cn");

export default function Layout() {
  return (
    <ConfigProvider locale={locale}>
      <Outlet />
    </ConfigProvider>
  );
}
