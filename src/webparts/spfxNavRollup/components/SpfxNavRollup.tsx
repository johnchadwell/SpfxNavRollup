import * as React from 'react';
import styles from './SpfxNavRollup.module.scss';
import type { ISpfxNavRollupProps } from './ISpfxNavRollupProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';  
import _ from 'lodash';

// interface for the columns returned for a single row
export interface INavResults {
  Title: string;
  Path: string;
  WebId: string;
}

// web part state containing array of INavResults
export interface ISpfxNavRollupState{    
  navResults: INavResults[];
}    

// interface that will allow an "any" array for nav results. The nav results will return as:
// response.PrimaryQueryResult.RelevantResults.Table.Rows - difficult to assign typescript type.
// Allows for Map function to work on a type any[].
export interface INavRollupResults {
  results: any[];
}

export default class SpfxNavRollup extends React.Component<ISpfxNavRollupProps, ISpfxNavRollupState> {

  public constructor(props: ISpfxNavRollupProps, state: ISpfxNavRollupState){    
    super(props);   
    this.state = {    
      // navResults: [],
      navResults: [
        {    
          Title: "",   
          Path: "",   
          WebId: "",
        }
      ]
    };    
  }    

  // private compare = function(a: string,b: string) {

  //   var pathA = a.toUpperCase();
  //   var pathB = b.toUpperCase();
  //   var pathA = r.result.Cells[1].Value.toUpperCase();
  //   // var pathB = b.Cells.results[6].Value.toUpperCase();
  //   var comparison = 0;
  //   if (pathA > pathB) {
  //     comparison = 1;
  //   } else if (pathA < pathB) {
  //     comparison = -1;
  //   }
  //   return comparison;
  // }

  public componentDidMount() {  

    console.log("componentDidMount");

    this.search().then((response) => {
      console.log("results: " + JSON.stringify(response.PrimaryQueryResult.RelevantResults.Table.Rows));
      var r = {} as INavRollupResults;
      r.results = response.PrimaryQueryResult.RelevantResults.Table.Rows;
      var all = [] as INavResults[];
      r.results.map((result, index) => {
        console.log("Title: " + result.Cells[0].Value);
        console.log("Path: " + result.Cells[1].Value);
        console.log("WebId: " + result.Cells[2].Value);
        var r = {} as INavResults;
        r.Title = result.Cells[0].Value;
        r.Path = result.Cells[1].Value;
        r.WebId = result.Cells[2].Value;
        all.push(r);
      })
      // doesn't work - why?
      // all.sort((a: INavResults, b: INavResults) => (a.Path as any) - (b.Path as any));
      console.log("all sorted: " + JSON.stringify(all));

      const allSorted: INavResults[] = _.sortBy(all, 'Path');
      console.log("allSorted: " + JSON.stringify(allSorted));
      this.setState({navResults : allSorted});      
    });
  }  


  
  private search(): Promise<any> {

    var restApi = "";
    if (this.props.QueryUrl===undefined || this.props.QueryUrl.length==0) {
      var url = "https://" + window.location.hostname;
      console.log("url: " + url)
      var restApi = url + "/_api/search/query?querytemplate='path:\"" + url + "\" contentclass:\"STS_Site\"'&selectproperties='Title,Path,WebId,SiteId,OriginalPath'";
    } else {
      var restApi = this.props.QueryUrl;
    }
    
    console.log("restApi = " + restApi);
    return this.props.context.spHttpClient.get(`${restApi}`, SPHttpClient.configurations.v1,
      {
        headers: [
          // default OData V4
          ['accept', 'application/json;odata.metadata=none']
          // OData V3
          // ['accept', 'application/json;odata=nometadata'],
          // ['odata-version', '']
        ]
      })
      .then((response: SPHttpClientResponse) => {
      if (response.ok) {
        return response.json();
      } else {
        console.log("Error: failed on URL " + restApi + ". Error = " + response.statusText);
        return null;
      }
    });
  
  }

  
  public render(): React.ReactElement<ISpfxNavRollupProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.spfxNavRollup} ${hasTeamsContext ? styles.teams : ''}`}>
        {/* <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>
        <div>
          <h3>Welcome to SharePoint Framework!</h3>
          <p>
            The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It&#39;s the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
          </p>
          <h4>Learn more about SPFx development:</h4>
          <ul className={styles.links}>
            <li><a href="https://aka.ms/spfx" target="_blank" rel="noreferrer">SharePoint Framework Overview</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank" rel="noreferrer">Use Microsoft Graph in your solution</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank" rel="noreferrer">Build for Microsoft Teams using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank" rel="noreferrer">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank" rel="noreferrer">Publish SharePoint Framework applications to the marketplace</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank" rel="noreferrer">SharePoint Framework API reference</a></li>
            <li><a href="https://aka.ms/m365pnp" target="_blank" rel="noreferrer">Microsoft 365 Developer Community</a></li>
          </ul>
        </div> */}

        <div>{this.props.description}</div>

        <div className={ styles.spfxNav }>

            {this.state.navResults.map(item => {    
                  
                  return (<div key={item.WebId}>    

                      <div style={{color: this.props.Color, background: this.props.Background, width: this.props.Width + 'px', height: this.props.Height + 'px', margin: this.props.Margin, padding: this.props.Padding, border: this.props.Border, borderRadius: this.props.BorderRadius, boxShadow: this.props.BoxShadow}}>
                      <a style={{ height:"100%", width:"100%", textDecoration: "none" }} href={item.Path} target="target">
                        <div style={{height:"100%", width:"100%", position: "relative", display: "table", textAlign: "left", fontFamily: this.props.Font}}>
                          <p style={{display: "table-cell", verticalAlign: "Top", color: this.props.Color}}>{item.Title}</p>
                        </div>
                      </a>
                      </div>              

                    </div>
                    );    
                })}    
  
      </div>

        {/* {this.state.navResults.map(function(item,key){  
                
          return (
          <div key={key}>  
              <div>{item.Title}</div>  
              <div>{item.Path}</div>  
              <div>{item.WebId}</div>
          </div>);  
        })}   */}

      </section>
    );
  }
}

