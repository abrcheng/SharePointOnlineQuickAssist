import { ICrawlLog,ICrawlLogs } from "./ICrawlLog";
import { DetailsList, DetailsListLayoutMode, SelectionMode, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import * as React from 'react';
import * as strings from 'SharePointOnlineQuickAssistWebPartStrings';
export default class CrawlLogGrid extends React.Component<ICrawlLogs>
{   
    private columns: IColumn[] = [   
        {
            key: 'TimeStamp',
            name: strings.SD_CrawlTimeStamp,
            fieldName: 'TimeStamp',
            minWidth: 100,
            maxWidth: 100,
            isResizable: true,
            isCollapsible: true,
            data: 'string',          
            onRender: (item: ICrawlLog) => {
            return <div style={{whiteSpace: 'pre-wrap', overflowWrap: 'break-word'}}>{item.TimeStamp}</div>;
            },
            isPadded: true,                 
        },    
        {
            key: 'Path',
            name: strings.SD_CrawlPath,
            fieldName: 'Path',
            minWidth: 120,
            maxWidth: 120,
            isResizable: true,
            isCollapsible: true,
            data: 'string',          
            onRender: (item: ICrawlLog) => {
              return <div style={{whiteSpace: 'pre-wrap', overflowWrap: 'break-word'}}>{item.FullUrl}</div>;
            },            
            isPadded: true                        
          },
          {
            key: 'ErrorCode',
            name: strings.SD_CrawlErrorCode,
            fieldName: 'ErrorCode',
            minWidth: 30,
            maxWidth: 60,
            isResizable: true,
            isCollapsible: true,
            data: 'string',          
            onRender: (item: ICrawlLog) => {
              return <span>{item.ErrorCode}</span>;
            },
            isPadded: true            
          },
            {
                key: 'IsDeleted',
                name: strings.SD_CrawlIsDeleted,
                fieldName: 'IsDeleted',
                minWidth: 30,
                maxWidth: 60,
                isResizable: true,
                isCollapsible: true,
                data: 'string',          
                onRender: (item: ICrawlLog) => {
                return <span>{item.IsDeleted}</span>;
                },
                isPadded: true                                     
            },
            {
                key: 'DeletePending',
                name: strings.SD_CrawlDeletePending,
                fieldName: 'DeletePending',
                minWidth: 30,
                maxWidth: 60,
                isResizable: true,
                isCollapsible: true,
                data: 'string',          
                onRender: (item: ICrawlLog) => {
                return <span>{item.DeletePending}</span>;
                },
                isPadded: true                                     
            },
            {
                key: 'DeleteReason',
                name: strings.SD_CrawlDeleteReason,
                fieldName: 'DeleteReason',
                minWidth: 30,
                maxWidth: 60,
                isResizable: true,
                isCollapsible: true,
                data: 'string',          
                onRender: (item: ICrawlLog) => {
                return <span>{item.DeleteReason}</span>;
                },
                isPadded: true                                     
            },
            {
                key: 'ExReason',
                name: strings.SD_CrawlExclusionReason,
                fieldName: 'ExReason',
                minWidth: 30,
                maxWidth: 60,
                isResizable: true,
                isCollapsible: true,
                data: 'string',          
                onRender: (item: ICrawlLog) => {
                return <span>{item.ExclusionReason}</span>;
                },
                isPadded: true                                     
            },
            {
                key: 'ErrorDesc',
                name: strings.SD_CrawlErrorDesc,
                fieldName: 'ErrorDesc',
                minWidth: 120,
                maxWidth: 120,
                isResizable: true,
                isCollapsible: true,
                data: 'string',          
                onRender: (item: ICrawlLog) => {
                return <div style={{whiteSpace: 'pre-wrap', overflowWrap: 'break-word'}}>{item.ErrorDesc}</div>;
                },
                isPadded: true                                     
            }  
      ];

      public state = {
        items:this.props.items,
        columns:this.columns};

        public render():React.ReactElement<ICrawlLogs>
        {       
            const { items, columns} = this.state;        
            return <div>
                 <div><span>Crawl Logs:</span></div>
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
}