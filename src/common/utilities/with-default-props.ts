import { ComponentType } from 'react';

const withDefaultProps = <P extends object, DP extends Partial<P> = Partial<P>>(
    defaultProps: DP,
    Cmp: ComponentType<P>
  ) => {
    type PropsExcludingDefaults = Pick<P,Exclude<keyof P, keyof DP>>;
    type RecomposedProps = Partial<DP> & PropsExcludingDefaults;
    Cmp.defaultProps = defaultProps;
    return (Cmp as ComponentType<any>) as ComponentType<RecomposedProps>;
  };
  
  export default withDefaultProps;