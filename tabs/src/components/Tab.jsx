import React from "react";
import { app, teamsCore } from "@microsoft/teams-js";
import './App.css';

function logMessage(action, message)  
{
  return "<span style='font-weight:bold'>" + action + "</span>" + message + "</br>";
}

class Tab extends React.Component {
  constructor(props){
    super(props)
    this.state = {
      context: {},
      logs: []
    }
  }

  //React lifecycle method that gets called once a component has finished mounting
  //Learn more: https://reactjs.org/docs/react-component.html#componentdidmount
  componentDidMount(){
    app.initialize().then(() => {
      // Get the user context from Teams and set it in the state
      app.getContext().then(async (context) => {
        app.notifySuccess();
        this.setState({
          logs: [...this.state.logs, logMessage("OnAppContext", "Notify success of loading - Loaded Teams context.")]
        });

        this.setState({
          meetingId: context.meeting.id,
          userName: context.user.userPrincipalName
        });

        // App caching is only available in meeting's side panel.
        if (context.page.frameContext === "sidePanel") {
          teamsCore.registerBeforeUnloadHandler((readyToUnload: any) => {
            this.setState({
              logs: [...this.state.logs, logMessage("BeforeUnload", "Sending readyToUnload to Teams...")]
            });
            readyToUnload();
            return true;
          });

          teamsCore.registerOnLoadHandler((context) => {
            this.setState({
              logs: [...this.state.logs, logMessage("OnLoadContext", "Entity Id: " + context.entityId + ", Content Url: " + context.contentUrl)]
            });

            app.notifySuccess();
          });

          this.setState({
            logs: [...this.state.logs, logMessage("Summary", "Registered load & beforeUnload handlers. Ready for app caching.")]
          });
        }
      });
    });
    // Next steps: Error handling using the error object
  }

  render() {
    let meetingId = this.state.meetingId ?? "";
    let userPrincipleName = this.state.userName ?? "";
    let logs = this.state.logs;
    return (
    <div>
      <h1>In-meeting app sample</h1>
      <h3>Principle Name:</h3>
      <p>{userPrincipleName}</p>
      <h3>Meeting ID:</h3>
      <p>{meetingId}</p> 
      <div>
        {logs.map((log) => {
          return <div dangerouslySetInnerHTML={{ __html: log }}></div>
        })}
      </div>
    </div>
    );
  }
}

export default Tab;