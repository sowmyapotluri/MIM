import * as React from 'react';
import * as microsoftTeams from "@microsoft/teams-js";
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { Input, Loader, Button, Flex, FlexItem, Text, Icon, Dropdown, Checkbox, Accordion, Avatar, Segment } from '@fluentui/react';
import { SearchIcon, FilesImageIcon } from '@fluentui/react-icons-northstar';
import { DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { Guid } from "guid-typescript";
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import { AxiosResponse } from "axios";
import "./Dashboard.scss";
import { isNullOrUndefined } from 'util';
import { ApplicationInsights, SeverityLevel } from '@microsoft/applicationinsights-web';
import { ReactPlugin, withAITracking } from '@microsoft/applicationinsights-react-js';
import { createBrowserHistory } from "history";
import { orderBy, forEach } from 'lodash';
import { getIncidents } from '../../apis/apiList';
import moment from 'moment';
let reactPlugin = new ReactPlugin();
const browserHistory = createBrowserHistory({ basename: '' });


export interface IDashboardState {
    newIncidents: IIncidentEntity[],
    suspendedIncidents: IIncidentEntity[],
    restoredIncidents: IIncidentEntity[],
    loader: boolean,
    masterData: IIncidentEntity[],
    selectedIncident: string
}

export interface IConferenceRooms {
    code: number,
    available: boolean,
    bridgeURL: string,
    channelId: string,
}

export interface IIncidentEntity {
    description: string,
    number: string,
    priority: string,
    shortDescription: string,
    status: string,
    createdOn: string,
    id: string,
    updatedOn: string,
    bridgeDetails: IConferenceRooms,
    linkToThread: string,
    currentActivity: string,
    assignedTo: IUser,
    bridgeId: number,
    bridgeLink: string,
    requestedBy: IUser
}

export interface IUser {
    id: string,
    userPrincipalName: string,
    displayName: string,
    teamsUserId: string,
    serviceUrl: string,
    profilePicture: string
}

export const Priority = {
    Low: 1,
    Normal: 2,
    High: 3
}

export default class Dashboard extends React.Component<{}, IDashboardState> {

    token?: string | null = null;
    telemetry: any = undefined;
    incidentNumber?: string | null = null;
    incidentId?: string | null = null;
    assignedToChanged?: boolean = false;
    // appInsights: ApplicationInsights;



    constructor(props: {}) {
        super(props);
        initializeIcons();
        this.state = {
            loader: false,
            newIncidents: [
                {
                    description: "",
                    number: "",
                    priority: "",
                    shortDescription: "",
                    status: "",
                    createdOn: "",
                    id: "",
                    updatedOn: "",
                    bridgeDetails: {
                        code: 0,
                        bridgeURL: "",
                        available: true,
                        channelId: ""
                    },
                    linkToThread: "",
                    currentActivity: "",
                    assignedTo: {
                        id: "",
                        userPrincipalName: "",
                        displayName: "",
                        teamsUserId: "",
                        serviceUrl: "",
                        profilePicture: ""
                    },
                    bridgeId: 0,
                    bridgeLink: "",
                    requestedBy: {
                        id: "",
                        userPrincipalName: "",
                        displayName: "",
                        teamsUserId: "",
                        serviceUrl: "",
                        profilePicture: ""
                    }
                }
            ],
            restoredIncidents: [],
            suspendedIncidents: [],
            masterData: [],
            selectedIncident: ""
        }
        let search = window.location.search;
        let params = new URLSearchParams(search);
        this.telemetry = params.get("telemetry");
        this.token = params.get("token");
        // this.appInsights = new ApplicationInsights({
        //     config: {
        //         instrumentationKey: this.telemetry,
        //         extensions: [reactPlugin],
        //         extensionConfig: {
        //             [reactPlugin.identifier]: { history: browserHistory }
        //         }
        //     }
        // });
        // this.appInsights.loadAppInsights();

    };

    public componentDidMount = () => {
        microsoftTeams.initialize();
        microsoftTeams.getContext((context) => {
            console.log("microsoft teams", context)
        });
        document.removeEventListener("keydown", this.escFunction, false);
        this.getIncidents123();
    }

    public componentWillUnmount = () => {
        document.removeEventListener("keydown", this.escFunction, false);
    }

    private escFunction = (e: KeyboardEvent) => {
        if (e.keyCode === 27 || (e.key === "Escape")) {
            microsoftTeams.tasks.submitTask({ "output": "failure" });
        }
    }

    private getIncidents123 = async () => {

        // this.setState({
        //     loader: true
        // });

        // getIncidents(moment().format("YYYY/MM/DD"))
        await fetch("/api/IncidentApi/GetAllIncidents?weekDay=" + moment().format("YYYY/MM/DD"), {
            method: "GET",
            headers: {
                "Content-Type": "application/json",
                "Authorization": "Bearer " + this.token
            },
        })
            .then(async (res: any) => {
                if (res.status === 401) {
                    const response = res.data;
                    if (response) {
                        // this.setState({
                        //     errorResponseDetail: {
                        //         errorMessage: response.message,
                        //         statusCode: response.code,
                        //     }
                        // })
                    }

                    // this.setState({ authorized: false });
                    // this.appInsights.trackTrace({ message: `User ${this.userObjectIdentifier} is unauthorized!`, severityLevel: SeverityLevel.Warning });
                    return response;
                }
                else if (res.status === 200) {
                    let response = await res.json(); //res.data;
                    let incidents: IIncidentEntity[] = [];
                    for (let i = 0; i < response.length; i++) {
                        let incident: IIncidentEntity = {
                            // bridge: response[i].Id,
                            description: response[i].Description,
                            number: response[i].Number,
                            priority: response[i].Id,
                            shortDescription: response[i].ShortDescription,
                            status: response[i].Status,
                            createdOn: response[i].CreatedOn,
                            id: response[i].RowKey,
                            updatedOn: response[i].UpdatedOn,
                            bridgeDetails: {
                                code: response[i].BridgeId,
                                bridgeURL: response[i].BridgeLink,
                                available: true,
                                channelId: ""
                            },
                            linkToThread: response[i].TeamConversationId,
                            currentActivity: response[i].CurrentActivity,
                            assignedTo: {
                                displayName: response[i].AssignedTo.displayName,
                                id: response[i].AssignedTo.id,
                                profilePicture: response[i].AssignedTo.profilePicture,
                                serviceUrl: "",
                                teamsUserId: "",
                                userPrincipalName: ""
                            },
                            bridgeId: response[i].BridgeId,
                            bridgeLink: response[i].BridgeLink,
                            requestedBy: {
                                displayName: response[i].RequestedBy.displayName,
                                id: response[i].RequestedBy.id,
                                profilePicture: response[i].RequestedBy.profilePicture,
                                serviceUrl: "",
                                teamsUserId: "",
                                userPrincipalName: ""
                            }
                        }
                        incidents.push(incident);
                    }
                    console.log("allIncidents", incidents)
                    this.setState({
                        newIncidents: incidents.filter(incident => incident.status === "1"),
                        suspendedIncidents: incidents.filter(incident => incident.status === "2"),
                        restoredIncidents: incidents.filter(incident => incident.status === "3"),
                        masterData: incidents,
                        loader: false
                    });
                }
                else {
                    // this.setMessage(this.state.resourceStrings.ExceptionResponse, Constants.ErrorMessageRedColor, false);
                    // this.appInsights.trackTrace({ message: `'SearchRoomAsync' - Request failed:${res.status}`, severityLevel: SeverityLevel.Warning });
                }

            });
    }

    private showDetails = (id: string, flag: boolean) => {
        if (flag) {
            this.setState({
                selectedIncident: id
            });
        }
        else {
            this.setState({
                selectedIncident: ""
            });
        }
    }

    private deeplinkToThread = (link: string) => {
        console.log("link",link)
        microsoftTeams.executeDeepLink(link);
        // microsoftTeams.tasks.submitTask({"deeplinkExecuted":"success"});
    }

    private renderBody = (incidents: IIncidentEntity[]) => {
        let rows: any[] = [];
        let requestedUser, assignTo;
        (incidents.map((incident: IIncidentEntity, index: number) => {
            if (typeof (incident.requestedBy.profilePicture) == 'string' && incident.requestedBy.profilePicture.indexOf('error') < 0) {
                requestedUser = (
                    // <Segment >
                    <Flex gap="gap.small">
                        <FlexItem push>
                            <Avatar image={`data:image/jpeg;base64,${incident.requestedBy.profilePicture}`} />
                        </FlexItem>
                        <FlexItem grow>
                            <Text className="userPadding" content={incident.requestedBy.displayName} />
                        </FlexItem>
                    </Flex>
                    // </Segment>
                )
            } else {
                requestedUser = (
                    // <Segment >
                    <Flex gap="gap.small">
                        <FlexItem push>
                            <Avatar name={incident.requestedBy.displayName} />
                        </FlexItem>
                        <FlexItem grow>
                            <Text className="userPadding" content={incident.requestedBy.displayName} />
                        </FlexItem>
                    </Flex>
                    // </Segment>
                )
            }
            if (typeof (incident.assignedTo.profilePicture) == 'string' && incident.assignedTo.profilePicture.indexOf('error') < 0) {
                assignTo = (
                    // <Segment >
                    <Flex gap="gap.small">
                        <FlexItem push>
                            <Avatar image={`data:image/jpeg;base64,${incident.assignedTo.profilePicture}`} />
                        </FlexItem>
                        <FlexItem grow>
                            <Text className="userPadding" content={incident.assignedTo.displayName} />
                        </FlexItem>
                    </Flex>
                    // </Segment>
                )
            } else {
                assignTo = (
                    // <Segment >
                    <Flex gap="gap.small">
                        <FlexItem push>
                            <Avatar name={incident.assignedTo.displayName} />
                        </FlexItem>
                        <FlexItem grow>
                            <Text className="userPadding" content={incident.assignedTo.displayName} />
                        </FlexItem>
                    </Flex>
                    // </Segment>
                )
            }
            rows.push(
                <tr onMouseLeave={() => { this.showDetails(incident.id, false) }} onMouseEnter={() => { this.showDetails(incident.id, true) }}>
                    <td>
                        <Text key={"number" + index} content={incident.number} />
                    </td>
                    <td>
                        <Text key={"description" + index} content={incident.shortDescription} />
                    </td>
                    <td>
                        <Text key={"brideId" + index} content={incident.bridgeId} color="brand" weight="semibold" />
                    </td>
                    <td>
                        <Text key={"state" + index} content={incident.status} color="brand" weight="semibold" />
                    </td>
                    <td>
                        <Text key={"createdOn" + index} content={incident.createdOn} color="brand" weight="semibold" />
                    </td>
                    <td>
                        <Text key={"updatedOn" + index} content={incident.updatedOn} color="brand" weight="semibold" />
                    </td>
                    <td>
                        <div onClick={()=>this.deeplinkToThread(incident.linkToThread)}>
                            <Flex gap="gap.small">
                                <FlexItem push>
                                <Avatar image={`data:image/jpeg;base64,${incident.assignedTo.profilePicture}`} />
                                </FlexItem>
                                <FlexItem grow>
                                    <Text className="userPadding" content={"BART"} />
                                </FlexItem>
                            </Flex>

                        </div>
                    </td>

                    {/* <div hidden={index !== this.state.workstreams.length - 1}>
                    <Flex gap="gap.smaller">
                        <Icon iconName="add" className="addIcon" />
                        <Button text content="Add another workstream" onClick={this.addWorkstreams} />
                    </Flex>
                </div> */}

                </tr>
            );
            rows.push(
                <tr hidden={!(this.state.selectedIncident === incident.id)}>
                    <td>
                        <Text weight="semibold" content="Requested By" />
                    </td>
                    <td colSpan={2}>
                        {requestedUser}
                    </td>
                    <td>
                        <Text weight="semibold" content="Assigned To" />
                    </td>
                    <td colSpan={2}>
                        {assignTo}
                    </td>
                    <td>
                        <Text weight="semibold" content={"Current Activity: " + incident.currentActivity} />
                    </td>
                </tr>
            )
            // return (
            //     <div>


            //     </div>
            // )
        }))
        return (

            <table className="table table-borderless">
                <thead>
                    <tr>
                        <th>Incident</th>
                        <th>Description</th>
                        <th>Conference Line</th>
                        <th>State</th>
                        <th>Created on</th>
                        <th>Last updated</th>
                        <th>Go to channel thread</th>
                    </tr>
                </thead>
                <tbody>
                    {rows}
                </tbody>
            </table>
        )
    }

    private searchIncidents = (e: React.SyntheticEvent<HTMLElement, Event>) => {
        let searchQuery = (e.target as HTMLInputElement).value;
        if (!searchQuery) // If Search text cleared
        {
            this.setState({
                newIncidents: this.state.masterData.filter(incident => incident.status === "1"),
                suspendedIncidents: this.state.masterData.filter(incident => incident.status === "2"),
                restoredIncidents: this.state.masterData.filter(incident => incident.status === "3"),
            })
        }
        else {
            this.setState({
                newIncidents: this.state.masterData.filter((x: IIncidentEntity) => x.shortDescription.toLowerCase().includes(searchQuery.toLowerCase())),
                suspendedIncidents: this.state.masterData.filter((x: IIncidentEntity) => x.shortDescription.toLowerCase().includes(searchQuery.toLowerCase())),
                restoredIncidents: this.state.masterData.filter((x: IIncidentEntity) => x.shortDescription.toLowerCase().includes(searchQuery.toLowerCase()))
            })
        }
    }


    public render(): JSX.Element {

        const panels = [
            {
                key: 'new',
                title: 'New',
                content: (
                    this.renderBody(this.state.newIncidents)
                ),
            },
            {
                key: 'suspended',
                title: 'Suspended',
                content: (
                    this.renderBody(this.state.suspendedIncidents)
                ),
            },
            {
                key: 'restored',
                title: "Service Restored",
                content: (
                    this.renderBody(this.state.restoredIncidents)
                ),
            },
        ]

        if (this.state.loader) {
            return (
                <div className="emptyContent">
                    <Loader />
                </div>
            );
        }
        else {
            return (
                <div className="taskModule">
                    <div className="formContainer ">
                        <Flex>
                            <FlexItem>
                                <Text size="largest" content="Due this week" />
                            </FlexItem>
                            <FlexItem push>
                                <Input icon={<SearchIcon />} placeholder="Search" onChange={this.searchIncidents} />
                            </FlexItem>
                        </Flex>
                        <Accordion panels={panels} exclusive />

                        {/* <div hidden={index !== this.state.workstreams.length - 1}> */}

                        {/* </div> */}
                    </div>
                </div >
            );
        }
    }
}