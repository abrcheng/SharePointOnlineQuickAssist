import { DetailsList, DetailsListLayoutMode, SelectionMode, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import {  
    DefaultButton,
    TextField,   
  } from 'office-ui-fabric-react/lib/index';
import * as React from 'react';
import * as strings from 'SharePointOnlineQuickAssistWebPartStrings';
import { ICrawledProperty,ICrawledProperties } from "./ICrawledProperty";
import SPOQAHelper from '../../../Helpers/SPOQAHelper';
export default class CrawledPropertyGrid extends React.Component<ICrawledProperties>
{   
    private columns: IColumn[] = [   
        {
            key: 'Name',
            name: strings.SD_PropertyName,
            fieldName: 'Name',
            minWidth: 200,
            maxWidth: 400,
            isResizable: true,
            isCollapsible: true,
            data: 'string',          
            onRender: (item: ICrawledProperty) => {
                return <div style={{whiteSpace: 'pre-wrap', overflowWrap: 'break-word'}}>{item.Name}</div>;
            },
            isPadded: true,                 
        }]; 

          public state = {
            items:this.props.items,
            columns:this.columns             
        };
    
            public render():React.ReactElement<ICrawledProperties>
            {       
                const { items, columns} = this.state;        
                return <div>
                            <div>
                                <span>
                                    <TextField
                                        label={strings.SD_PropertyFilter}
                                        multiline={false}
                                        onChange={(e)=>{let text:any = e.target; this.FilterCrawledProperty(text.value);}}
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
                // Export CrawledProperty
                SPOQAHelper.JSONToCSVConvertor(this.props.items, true, "CrawledProperty");
            }
            
            private FilterCrawledProperty(cpFilter:string) {
                cpFilter = cpFilter.trim();
                if(cpFilter.length >=1)
                {
                    var filteredItems = [];
                    this.props.items.forEach(e=>{
                        if((e.Name && e.Name.toLowerCase().indexOf(cpFilter.toLowerCase()) >=0))
                        {
                            filteredItems.push(
                                {
                                    Name:e.Name                                    
                                }
                            );
                        }
                    });
                    this.setState({items:filteredItems});
                }
                else if(!cpFilter || cpFilter=="")
                {
                    this.setState({items:this.props.items});
                }
            }
}