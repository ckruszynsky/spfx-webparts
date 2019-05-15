
/**
 * Creates a closure and stores/infers defaultProps type information via generic parameter.
 * This should be used to get default props for components that use generic Props. For example a button 
 * component : Button<T> Component<Props<T>>
 * for more information read : https://medium.com/@martin_hotell/react-typescript-and-defaultprops-dilemma-ca7f81c661c7
 * @param defaultProps 
 */
export const createPropsGetter = <DP extends object>(defaultProps: DP) => {
    return <P extends Partial<DP>>(props:P) => {
        //we are extracting default props from component props api type
        type PropsExcludingDefaults = Pick<P, Exclude<keyof P, keyof DP>>;

        //we are re-creating our props definition by creating an intersection type
        //between Props without Defaults and NonNullable DefaultProps
        type RecomposedProps = DP & PropsExcludingDefaults;

        //we are returning the same props that we got as argument - identity function
        //also, we are turning off compiler and casting the type to our recomposed type
        return (props as any) as RecomposedProps;
    };
};
