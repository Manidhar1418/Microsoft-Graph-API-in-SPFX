/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
import * as React from 'react';
import styles from './Meetingattendence.module.scss';
import type { IMeetingattendenceProps } from './IMeetingattendenceProps';
// import { escape } from '@microsoft/sp-lodash-subset';
import { IPeoplePickerContext, PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import "@pnp/sp/fields";
import "@pnp/sp/webs";
import "@pnp/sp/items";
import "@pnp/sp/folders";
import "@pnp/sp/lists";
import "@pnp/sp/site-groups/web";
import "@pnp/sp/files";
import "@pnp/sp/profiles";
import "@pnp/sp/site-groups";
import { ComboBox } from '@fluentui/react';
import type { IComboBoxOption, IComboBoxStyles } from '@fluentui/react';
import {
  DatePicker,
  defaultDatePickerStrings,
} from '@fluentui/react';
import { Stack, IStackTokens } from '@fluentui/react';
import {  PrimaryButton } from '@fluentui/react/lib/Button';





  export interface IButtonExampleProps {
    // These are set based on the toggles shown above the examples (not needed in real code)
    disabled?: boolean;
    checked?: boolean;
  }

  // export const ButtonDefaultExample: React.FunctionComponent<IButtonExampleProps> = props => {
  //   const { disabled, checked } = props;
  




const options: IComboBoxOption[] = [
  { key: 'black', text: 'Black' },
  { key: 'blue', text: 'Blue' },
  { key: 'brown', text: 'Brown' },
  { key: 'cyan', text: 'Cyan' },
  { key: 'green', text: 'Green' },
  { key: 'mauve', text: 'Mauve' },
  { key: 'orange', text: 'Orange' },
  { key: 'pink', text: 'Pink' },
  { key: 'purple', text: 'Purple' },
  { key: 'red', text: 'Red' },
  { key: 'rose', text: 'Rose' },
  { key: 'violet', text: 'Violet' },
  { key: 'white', text: 'White' },
  { key: 'yellow', text: 'Yellow' },
];

const comboBoxStyles: Partial<IComboBoxStyles> = { root: { maxWidth: 300 } };
const stackTokens: IStackTokens = { childrenGap: 40 };



export default class Meetingattendence extends React.Component<IMeetingattendenceProps, {}> {
  private _peopplePicker:IPeoplePickerContext;
  constructor(props: IMeetingattendenceProps) {
    super(props);
   
   
 
    this._peopplePicker={
      absoluteUrl: this.props.context.pageContext.web.absoluteUrl,
      msGraphClientFactory:this.props.context.msGraphClientFactory,
      
      spHttpClient: this.props.context.spHttpClient
    }
   
  }

  private _getPeoplePickerItems(items: any[]) {
    console.log('Items:', items);
    
    console.log( this._peopplePicker.msGraphClientFactory)

  }
 

 


  
  public render(): React.ReactElement<IMeetingattendenceProps> {
    const {
      
      hasTeamsContext,

    } = this.props;
   
    console.log( this._peopplePicker.msGraphClientFactory)

    // eslint-disable-next-line @typescript-eslint/no-unused-vars

    return (
      <section className={`${styles.meetingattendence} ${hasTeamsContext ? styles.teams : ''}`}>
           <div style={{width:"300px"}}>
                <PeoplePicker
                  context={this._peopplePicker}
                  titleText="Host Name"
                  personSelectionLimit={1}
                  groupName={""} 
                  showtooltip={true}
                
                  disabled={false}
                  ensureUser={true}
                  onChange={this._getPeoplePickerItems}
                  
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000}
                />
              </div>
               
                

                <ComboBox
                label="Meetings Name"
                options={options}
                styles={comboBoxStyles}
                allowFreeInput
                autoComplete="on"
              />
              <div style={{width:"300px"}}>
              <DatePicker
              label='Date'
              placeholder="Select a date..."
              ariaLabel="Select a date"
              strings={defaultDatePickerStrings}
              disabled={true}
            />
            </div>
            <Stack horizontal tokens={stackTokens}>
      {/* <DefaultButton text="Standard" onClick={_alertClicked}  /> */}
      <PrimaryButton text="Primary" onClick={_alertClicked}  />
    </Stack>
 

      </section>
    );
  }
}
function _alertClicked(): void {
  alert('Clicked');
}
