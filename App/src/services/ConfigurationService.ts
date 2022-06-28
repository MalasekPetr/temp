import {
    HttpClientResponse,
    IHttpClientOptions
  } from '@microsoft/sp-http';
import { isUndefined } from 'lodash';
import { IConfiguration, IMailBoxApp } from '../models';
import { HttpService } from './../services';
export class ConfigurationService {
  constructor(private endpointUri: string) { }

  public addOrUpdateMailBoxApp(mailboxapp: IMailBoxApp): Promise<IMailBoxApp[]> {
    return new Promise<IMailBoxApp[]>((resolve, reject) => {
      const endpoint = `${this.endpointUri}/api/config`;
      const headers: Headers = new Headers();
      headers.append('Content-Type', 'application/json; charset=UTF-8');
      let options: IHttpClientOptions = { headers: headers, mode: 'cors' };
      options = { headers: headers, mode: 'cors', body: JSON.stringify(mailboxapp) };
      this.getMailBoxApp(mailboxapp.address)
        .then((res: HttpClientResponse) => {
          if (isUndefined(res.status)) {
            HttpService.Put(endpoint, options)
            .then((rawResponse: HttpClientResponse) => {
              return rawResponse.json();
            })
            .then((jsonResponse: IConfiguration) => {
              resolve(jsonResponse.mailBoxApps);
            })
            .catch((err: Record<string, unknown>) => {
              const response: Record<string, unknown>  = err;
              console.error(err);
              reject(response);
            });
          } else {
            HttpService.Post(endpoint, options)
            .then((rawResponse: HttpClientResponse) => {
              return rawResponse.json();
            })
            .then((jsonResponse: IConfiguration) => {
              resolve(jsonResponse.mailBoxApps);
            })
            .catch((err: Record<string, unknown>) => {
              const response: Record<string, unknown> = err;
              console.error(err);
              reject(response);
            });
          }
        });
    });
  }

  public getMailBoxApps(): Promise<IMailBoxApp[] | HttpClientResponse> {
    return new Promise<IMailBoxApp[]>((resolve, reject) => {
      const endpoint = `${this.endpointUri}/api/config`;
      const headers: Headers = new Headers();
      const options: IHttpClientOptions = { headers: headers, mode: 'cors' };
      HttpService.Get(endpoint, options)
        .then((rawResponse: HttpClientResponse) => {
          return rawResponse.json();
        })
        .then((jsonResponse: IConfiguration) => {
          // console.log(jsonResponse.mailBoxApps)
          resolve(jsonResponse.mailBoxApps);
        })
        .catch((err: Record<string, unknown>) => {
          const response: Record<string, unknown> = err;
          console.error(err);
          reject(response);
        });
    });
  }

  public getMailBoxApp(address: string): Promise<IMailBoxApp | HttpClientResponse> {
    return new Promise<IMailBoxApp>((resolve, reject) => {
      const endpoint = `${this.endpointUri}/api/config/${address}`;
      const headers: Headers = new Headers();
      const options: IHttpClientOptions = { headers: headers, mode: 'cors' };
      HttpService.Get(endpoint, options)
        .then((rawResponse: HttpClientResponse) => {
          return rawResponse.json();
        })
        .then((jsonResponse: IMailBoxApp) => {
          // console.log(jsonResponse)
          resolve(jsonResponse);
        })
        .catch((err: Record<string, unknown>) => {
          const response: Record<string, unknown> = err;
          console.error(err);
          reject(response);
        });
    });
  }
}