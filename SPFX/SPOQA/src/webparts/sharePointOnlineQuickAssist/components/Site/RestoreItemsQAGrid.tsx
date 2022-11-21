import { IRestoreItem,IRestoreItems } from "./IRestoreItem";
import { DetailsList, DetailsListLayoutMode, ColumnActionsMode, SelectionMode, IColumn,IDetailsHeaderProps } from 'office-ui-fabric-react/lib/DetailsList';
import {ContextualMenu, IContextualMenuProps, IContextualMenuItem, DirectionalHint} from 'office-ui-fabric-react/lib/ContextualMenu';
import * as React from 'react';
import * as strings from 'SharePointOnlineQuickAssistWebPartStrings';

export default class RestoreItemsQAGrid extends React.Component<IRestoreItems>
{   
    // private contextualMenuProps?: IContextualMenuProps;
    private onColumnClick = (event: React.MouseEvent<HTMLElement>, column: IColumn): void => {
      let isSortedDescending = column.isSortedDescending;
      // If we've sorted this column, flip it.
        if (column.isSorted) {
          isSortedDescending = !isSortedDescending;
        }
        this.sortByColumn(column, isSortedDescending);
    }
    
    private sortByColumn(column: IColumn, isSortedDescending:boolean)
    {
      const { columns } = this.state;
      let { items } = this.state;
        
       // Sort the items.
       items = _copyAndSort(items, column.fieldName!, isSortedDescending);

        // Reset the items and columns to match the state.
        this.setState({
          items: items,
          columns: columns.map(col => {
            col.isSorted = col.key === column.key;
            if (col.isSorted) {
              col.isSortedDescending = isSortedDescending;
            }
            return col;
          }),
          contextualMenuProps: undefined,
        });
    }
  
    private columns: IColumn[] = [       
        {
            key: 'Path',
            name: strings.RI_Path,
            fieldName: 'Path',
            minWidth: 200,
            maxWidth: 330,
            isResizable: true,
            isCollapsible: true,
            data: 'string',          
            onRender: (item: IRestoreItem) => {
              return <span>{item.Path}</span>;
            },            
            isPadded: true,
            onColumnClick: this.onColumnClick
          },
          {
            key: 'DeletedDate',
            name: strings.RI_DeletedDate,
            fieldName: 'DeletedDate',
            minWidth: 100,
            maxWidth: 140,
            isResizable: true,
            isCollapsible: true,
            data: 'string',          
            onRender: (item: IRestoreItem) => {
              return <span>{item.DeletedDate}</span>;
            },
            isPadded: true,
            onColumnClick: this.onColumnClick
          },
          {
                key: 'DeletedByEmail',
                name: strings.RI_DeletedByEmail,
                fieldName: 'DeletedByEmail',
                minWidth: 120,
                maxWidth: 180,
                isResizable: true,
                isCollapsible: true,
                data: 'string',          
                onRender: (item: IRestoreItem) => {
                return <span>{item.DeletedByEmail}</span>;
                },
                isPadded: true,
                onColumnClick: this.onColumnClick
            },
            {
                key: 'DeletedByName',
                name: strings.RI_DeletedBy,
                fieldName: 'DeletedByName',
                minWidth: 100,
                maxWidth: 120,
                isResizable: true,
                isCollapsible: true,
                data: 'string',          
                onRender: (item: IRestoreItem) => {
                return <span>{item.DeletedByName}</span>;
                },
                isPadded: true,
                onColumnContextMenu: (column, ev) => {
                  this.onColumnContextMenu(column, ev);},
                onColumnClick: (ev, column) => {
                  this.onColumnContextMenu(column, ev);},
              columnActionsMode: ColumnActionsMode.hasDropdown,                      
            },
            {
              key: 'Existing',
              name: strings.RI_Existing,
              fieldName: 'Existing',
              minWidth: 100,
              maxWidth: 120,
              isResizable: true,
              isCollapsible: true,
              data: 'string',          
              onRender: (item: IRestoreItem) => {
              return <span>{item.Existing.toString()}</span>;
              },
              isPadded: true,
              onColumnClick: this.onColumnClick,                      
          }           
      ];
     
    public state = {
        items:this.props.items,
        columns:this.columns,
        contextualMenuProps: undefined,
        filterByDeleteName:undefined,
        deleteNames:this.BuildDeleteNames(this.props.items)};

    public render():React.ReactElement<IRestoreItems>
    {       
        const { items, columns} = this.state;        
        return <div>
             {this.state.items && this.state.items.length >0 ?
              <DetailsList
              items={items}            
              columns={columns}
              selectionMode={SelectionMode.none}            
              layoutMode={DetailsListLayoutMode.fixedColumns}
              isHeaderVisible={true}                            
            />: strings.RI_NoData}    
             {this.state.contextualMenuProps && <ContextualMenu {...this.state.contextualMenuProps} />}        
        </div>;
    }
    
    private BuildDeleteNames(items:any[]):any[]
    {
        var deleteNames = [];
        for(var index=0; index< items.length; index++)
        {   
            var matched = false;
            deleteNames.forEach((delName)=>{
              if(delName.key == items[index].DeletedByName)
              {
                delName.count++;
                matched = true;
              }
            });

            if(!matched)
            {
              deleteNames.push({key:items[index].DeletedByName, count:1});
            }
        }

        return deleteNames;
    }

    private onContextualMenuDismissed = (): void => {
      this.setState({
          contextualMenuProps: undefined,
      });
    }
    
    private onColumnContextMenu = (column: IColumn, ev: React.MouseEvent<HTMLElement>): void => {
      if (column.columnActionsMode !== ColumnActionsMode.disabled) {
          this.setState({
              contextualMenuProps: this.getContextualMenuProps(ev, column),
          });
      }
    }

    private getContextualMenuProps = (ev: React.MouseEvent<HTMLElement>, column: IColumn): IContextualMenuProps => {
      // build sub menu
      const { deleteNames, filterByDeleteName } = this.state;  
      let subItems:IContextualMenuItem[] = [];
      deleteNames.forEach((delName)=>{
        let subItem:IContextualMenuItem={
          key:delName.key,
          name:`${delName.key} (${delName.count})`,
          canCheck: true,
          checked:delName.key == filterByDeleteName,
          onClick:(env, item) =>{this.filterByDeleteName(item);}
        };
        subItems.push(subItem);
      });
      const items: IContextualMenuItem[] = [
          {
              key: 'aToZ',
              name: 'A to Z',
              iconProps: { iconName: 'SortUp' },
              canCheck: true,
              checked: column.isSorted && !column.isSortedDescending,
              onClick:(env, item) =>{this.sortByColumn(column, false);}
          },
          {
              key: 'zToA',
              name: 'Z to A',
              iconProps: { iconName: 'SortDown' },
              canCheck: true,
              checked: column.isSorted && column.isSortedDescending,
              onClick:(env, item) =>{this.sortByColumn(column, true);}
          },
          {
            key: 'Filter',
            name: 'Filter',
            iconProps: { iconName: 'Filter' },
            canCheck: true,
            checked: this.state.filterByDeleteName,
            subMenuProps:{items:subItems}
        }
      ];
  
      return {
          items: items,
          target: ev.currentTarget as HTMLElement,
          directionalHint: DirectionalHint.bottomLeftEdge,
          gapSpace: 0,
          isBeakVisible: true,
          onDismiss: this.onContextualMenuDismissed,
      };
  }

  private filterByDeleteName(delNameItem:IContextualMenuItem)
  {
    const { columns } = this.state; // clean sort
    var cleanedSortColumns = columns.map(col => {
            col.isSorted=false;
            return col;
          });
     if(delNameItem.checked) // remove fitler 
     {
         this.setState(
           {items:this.props.items, filterByDeleteName:undefined, columns:cleanedSortColumns});
     }
     else // filter by key
     {
      this.setState(
        {items:this.props.items.filter(item=>item.DeletedByName == delNameItem.key)
          ,filterByDeleteName:delNameItem.key,
          columns:cleanedSortColumns});
     }
  }
}

function _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
  const key = columnKey as keyof T;
  return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
}
