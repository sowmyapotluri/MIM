import axios from './axiosJWTDecorator';
import { getBaseUrl } from '../configVariables';
import { AxiosResponse } from "axios";


let baseAxiosUrl = getBaseUrl() + '/api';

export const getAuthenticationMetadata = async (windowLocationOriginDomain: string, loginHint: string): Promise<AxiosResponse<string>> => {
    let url = `${baseAxiosUrl}/authenticationMetadata/GetAuthenticationUrlWithConfiguration?windowLocationOriginDomain=${windowLocationOriginDomain}&loginhint=${loginHint}`;
    return await axios.get(url, undefined, false);
}

export const getClientId = async (): Promise<AxiosResponse<string>> => {
    let url = baseAxiosUrl + "/authenticationMetadata/getClientId";
    return await axios.get(url);
}


export const GetResourceStringsApi = async (): Promise<AxiosResponse<JSON>> => {
    let url = baseAxiosUrl + "/ResourcesApi/GetResourceStrings";
    return await axios.get(url);
}

export const GetResourceStringstest = async (): Promise<AxiosResponse<string>> => {
    let url = baseAxiosUrl + "/ResourcesApi/GetResourceStringsTest";
    return await axios.get(url);
}
export const getSupportedTimeZones = async (): Promise<AxiosResponse<JSON>> => {
    let url = baseAxiosUrl + "/MeetingApi/GetSupportedTimeZonesAsync";
    return await axios.get(url);
}
export const getSavedTimeZone = async (): Promise<AxiosResponse<JSON>> => {
    let url = baseAxiosUrl + "/MeetingApi/GetUserTimeZoneAsync";
    return await axios.get(url);
}

export const saveUserTimeZone = async (payload: {}): Promise<AxiosResponse<void>> => {
    let url = baseAxiosUrl + "/MeetingApi/SaveTimeZoneAsyn";
    return await axios.post(url, payload);
}

export const getTopNRooms = async (payload: {}): Promise<AxiosResponse<void>> => {
    let url = baseAxiosUrl + "/MeetingApi/TopNRoomsAsync";
    return await axios.post(url, payload);
}

export const createMeeting= async (payload: {}): Promise<AxiosResponse<void>> => {
    let url = baseAxiosUrl + "/MeetingApi/CreateMeetingAsync";
    return await axios.post(url, payload);
}


export const filterRooms = async (payload: {}): Promise<AxiosResponse<void>> => {
    let url = baseAxiosUrl + "/MeetingApi/SearchRoomAsync";
    return await axios.post(url, payload);
}

