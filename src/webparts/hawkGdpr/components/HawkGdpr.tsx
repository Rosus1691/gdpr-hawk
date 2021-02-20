import * as React from 'react';
import styles from './HawkGdpr.module.scss';
import { IHawkGdprProps } from './IHawkGdprProps';
import { PivotTabsLargeExample } from './PivotLargeTabsExample';

import { sp } from "@pnp/sp-commonjs";





export default class HawkGdpr extends React.Component<IHawkGdprProps, {}> {
  

  constructor(props: IHawkGdprProps){
    super(props);
    sp.setup({
      spfxContext: this.props.context
    });
    
  }
 
  public render(): React.ReactElement<IHawkGdprProps> {
    
    
    return (
      <div className={ styles.hawkGdpr }>
        <div className={ styles.container }>
          {/* <PivotTabsLargeExample  context={this.props.context}/> */}
          <PivotTabsLargeExample sp={sp} />
         
        </div>
      </div>
    );
  }
}
