import axios from './axiosJWTDecorator';
import { getBaseUrl } from '../configVariables';
import { AxiosResponse } from "axios";
import { IIncidentEntity } from '../components/Dashboard/Dashboard';


let baseAxiosUrl = getBaseUrl() + '/api';

export const getAuthenticationMetadata = async (windowLocationOriginDomain: string, loginHint: string): Promise<AxiosResponse<string>> => {
    let url = `${baseAxiosUrl}/authenticationMetadata/GetAuthenticationUrlWithConfiguration?windowLocationOriginDomain=${windowLocationOriginDomain}&loginhint=${loginHint}`;
    return await axios.get(url, undefined, false);
}

export const getClientId = async (): Promise<AxiosResponse<string>> => {
    let url = baseAxiosUrl + "/authenticationMetadata/getClientId";
    return await axios.get(url);
}

export const getIncidents = async (date: string): Promise<AxiosResponse<IIncidentEntity[]>> => {
    let url = baseAxiosUrl + "/IncidentApi/GetAllIncidents?weekDay=" + date;
    return await axios.get(url);

}