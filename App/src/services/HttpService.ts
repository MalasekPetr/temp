import {
  HttpClient,
  IHttpClientOptions,
  HttpClientResponse
} from '@microsoft/sp-http';

export class HttpService {
  private static httpClient: HttpClient;

  public static onInit(httpClient: HttpClient): void {
    this.httpClient = httpClient;
  }

  public static async Get(url: string, options?: IHttpClientOptions): Promise<HttpClientResponse> {
    const response: HttpClientResponse = await this.httpClient.get(url, HttpClient.configurations.v1, options);
    return response;
  }

  public static async Post(url: string, options?: IHttpClientOptions): Promise<HttpClientResponse> {
    const response: HttpClientResponse = await this.httpClient.post(url, HttpClient.configurations.v1, options);
    return response;
  }

  public static async Put(url: string, options?: IHttpClientOptions): Promise<HttpClientResponse> {
    options.method = 'PUT';
    const response: HttpClientResponse = await this.httpClient.fetch(url, HttpClient.configurations.v1, options);
    return response;
  }
}