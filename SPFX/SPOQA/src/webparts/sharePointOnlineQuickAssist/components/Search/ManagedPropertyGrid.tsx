import { DetailsList, DetailsListLayoutMode, SelectionMode, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import {  
    DefaultButton,
    TextField,   
  } from 'office-ui-fabric-react/lib/index';
import * as React from 'react';
import * as strings from 'SharePointOnlineQuickAssistWebPartStrings';
import { IManagedProperty,IManagedProperties } from "./IManagedProperty";
import SPOQAHelper from '../../../Helpers/SPOQAHelper';
export default class ManagedPropertyGrid extends React.Component<IManagedProperties>
{   
    private columns: IColumn[] = [   
        {
            key: 'Name',
            name: strings.SD_PropertyName,
            fieldName: 'Name',
            minWidth: 100,
            maxWidth: 150,
            isResizable: true,
            isCollapsible: true,
            data: 'string',          
            onRender: (item: IManagedProperty) => {
                return <div style={{whiteSpace: 'pre-wrap', overflowWrap: 'break-word'}}>{item.Name}</div>;
            },
            isPadded: true,                 
        },    
        {
            key: 'Value',
            name: strings.SD_PropertyValue,
            fieldName: 'Value',
            minWidth: 200,
            maxWidth: 300,
            isResizable: true,
            isCollapsible: true,
            data: 'string',          
            onRender: (item: IManagedProperty) => {
              return <div style={{whiteSpace: 'pre-wrap', overflowWrap: 'break-word'}}>{item.Value}</div>;
            },            
            isPadded: true                        
          }]; 

          public state = {
            items:this.props.items,
            columns:this.columns             
        };
    
            public render():React.ReactElement<IManagedProperties>
            {       
                const { items, columns} = this.state;        
                return <div>
                            <div>
                                <span>
                                    <TextField
                                       label={strings.SD_PropertyFilter}
                                        multiline={false}
                                        onChange={(e)=>{let text:any = e.target; this.FilterManagedProperty(text.value);}}
                                        // value={this.state.keywordFilter}                                       
                                        // onKeyDown={(e)=>{if(e.keyCode ===13){}}}    
                                        style={{ display: 'inline'}}                                    
                                    />
                                     <DefaultButton
                                        text= {strings.RI_Export}
                                        style={{ display: 'inline', marginTop: '10px', marginLeft:"10px"}}
                                        onClick={() => {this.DoExport();}}
                                    />
                                </span>
                            </div>
                            {this.state.items && this.state.items.length >0 ?
                            <DetailsList
                            items={items}            
                            columns={columns}
                            selectionMode={SelectionMode.none}            
                            layoutMode={DetailsListLayoutMode.fixedColumns}                      
                            isHeaderVisible={true}                            
                            />: strings.RI_NoData} 
                </div>;
            }

            private DoExport():void
            {
                // Export ManagedProperty
                SPOQAHelper.JSONToCSVConvertor(this.props.items, true, "ManagedProperties");
            }
            
            private FilterManagedProperty(mpFilter:string) {
                mpFilter = mpFilter.trim();
                if(mpFilter.length >=1)
                {
                    var filteredItems = [];
                    this.props.items.forEach(e=>{
                        if((e.Name && e.Name.toLowerCase().indexOf(mpFilter.toLowerCase()) >=0)||(e.Value && e.Value.toLowerCase().indexOf(mpFilter.toLowerCase()) >=0))
                        {
                            filteredItems.push(
                                {
                                    Name:e.Name,
                                    Value:e.Value
                                }
                            );
                        }
                    });
                    this.setState({items:filteredItems});
                }
                else if(!mpFilter || mpFilter=="")
                {
                    this.setState({items:this.props.items});
                }
            }
}