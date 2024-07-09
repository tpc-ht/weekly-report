import { createStyles } from "antd-style";
const useStyles = createStyles(({ css, prefixCls }, { width }: { width: number }) => {
  return {
    container: css`
      & .${prefixCls}-form-item .${prefixCls}-form-item-label {
        width: ${width}px;
      }
    `,
  };
});
type FormLayout = {
  children: any;
  width?: number;
};
export default ({ children, width = 120 }: FormLayout) => {
  const { styles, cx, theme } = useStyles({ width });
  return <div className={styles.container}>{children}</div>;
};
