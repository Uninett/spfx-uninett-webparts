/**
 * Base communication layer for the SharePoint Framework
 *
 * @remarks
 * This package defines the base communication layer for
 * the SharePoint Framework.  For REST calls, it handles authentication,
 * logging, diagnostics, and batching.  It also simplifies requests by
 * adding default headers that follow the recommended best practices.
 *
 * @packagedocumentation
 */

import * as AuthenticationContext from 'adal-angular';
import { GraphRequestCallback } from '@microsoft/microsoft-graph-client';
import { Guid } from '@microsoft/sp-core-library';
import { Options } from '@microsoft/microsoft-graph-client';
import { ServiceKey } from '@microsoft/sp-core-library';
import { ServiceScope } from '@microsoft/sp-core-library';
import { SPEvent } from '@microsoft/sp-core-library';
import { SPEventArgs } from '@microsoft/sp-core-library';
import { URLComponents } from '@microsoft/microsoft-graph-client';

/* Excluded from this release type: _AadConstants */

/**
 * AadHttpClient is used to perform REST calls against an Azure AD Application.
 *
 * @remarks
 * For communicating with SharePoint, use the {@link SPHttpClient} class instead.
 * For communicating with Microsoft Graph, use the {@link MSGraphClient} class.
 *
 * @public
 * @sealed
 */
export declare class AadHttpClient {
    /**
     * The standard predefined AadHttpClientConfiguration objects for use with
     * the AadHttpClient class.
     */
    static readonly configurations: IAadHttpClientConfigurations;
    private static readonly _className;
    private _aadTokenConfiguration;
    private _aadTokenProvider;
    private _resourceUrl;
    private _serviceScope;
    private _fetchProvider;
    /**
     *
     * @param serviceScope - The service scope is needed to retrieve some of the class's internal components.
     * @param resourceUrl - The resource for which the token should be obtained.
     * @param options - Configuration options for the request to get an access token.
     *
     */
    constructor(serviceScope: ServiceScope, resourceEndpoint: string, options?: IAadHttpClientOptions);
    /**
     * Performs a REST service call.
     *
     * @remarks
     * Although the AadHttpClient subclass adds additional enhancements, the parameters and semantics
     * for HttpClient.fetch() are essentially the same as the WHATWG API standard that is documented here:
     * https://fetch.spec.whatwg.org/
     *
     * @param url - The endpoint URL that fetch will be called on.
     * @param configuration - Determines the default behavior of HttpClient; normally this should
     *   be the latest version number from HttpClientConfigurations.
     * @param options - Additional options that affect the request.
     * @returns A promise that will return the result.
     */
    fetch(url: string, configuration: AadHttpClientConfiguration, options: IHttpClientOptions): Promise<HttpClientResponse>;
    /**
     * Calls fetch(), but sets the method to "GET".
     *
     * @param url - The endpoint URL that fetch will be called on.
     * @param configuration - Determines the default behavior of HttpClient; normally this should
     *   be the latest version number from HttpClientConfigurations.
     * @param options - Additional options that affect the request.
     * @returns A promise that will return the result.
     */
    get(url: string, configuration: AadHttpClientConfiguration, options?: IHttpClientOptions): Promise<HttpClientResponse>;
    /**
     * Calls fetch(), but sets the method to "POST".
     *
     * @param url - The endpoint URL that fetch will be called on.
     * @param configuration - Determines the default behavior of HttpClient; normally this should
     *   be the latest version number from HttpClientConfigurations.
     * @param options - Additional options that affect the request.
     * @returns A promise that will return the result.
     */
    post(url: string, configuration: AadHttpClientConfiguration, options: IHttpClientOptions): Promise<HttpClientResponse>;
}

/**
 * Configuration for HttpClient.
 *
 * @remarks
 * The HttpClientConfiguration object provides a set of switches for enabling/disabling
 * various features of the HttpClient class.  Normally these switches are set
 * (e.g. when calling HttpClient.fetch()) by providing one of the predefined defaults
 * from HttpClientConfigurations, however switches can also be changed via the
 * HttpClientConfiguration.overrideWith() method.
 *
 * @public
 */
export declare class AadHttpClientConfiguration extends HttpClientConfiguration implements IAadHttpClientConfiguration {
    protected flags: IAadHttpClientConfiguration;
    /**
     * Constructs a new instance of HttpClientConfiguration with the specified flags.
     * The default values will be used for any flags that are missing or undefined.
     * If overrideFlags is specified, it takes precedence over flags.
     */
    constructor(flags: IAadHttpClientConfiguration, overrideFlags?: IAadHttpClientConfiguration);
    /**
     * @override
     */
    overrideWith(sourceFlags: IAadHttpClientConfiguration): AadHttpClientConfiguration;
}

/**
 * Returns a preinitialized version of the AadHttpClient for a given resource url.
 * For more information: {@link https://docs.microsoft.com/en-us/sharepoint/dev/spfx/use-aadhttpclient}
 *
 * @public
 */
export declare class AadHttpClientFactory {
    /**
     * The service key for AadHttpClientFactory.
     */
    static readonly serviceKey: ServiceKey<AadHttpClientFactory>;
    private _serviceScope;
    /* Excluded from this release type: __constructor */
    /**
     * Returns an instance of the AadHttpClient that communicates with the current tenant's configurable
     * Service Principal.
     * @param resourceEndpoint - The target AAD application's resource endpoint.
     */
    getClient(resourceEndpoint: string): Promise<AadHttpClient>;
    /* Excluded from this release type: _getStandardClient */
}

/**
 * This class allows a developer to obtain OAuth2 tokens from Azure AD.
 *
 * OAuth2 tokens are used to authenticate the user from the SharePoint page
 * to other services such as PowerBI, Sway, Exchange, Yammer, etc.
 *
 * @privateRemarks
 * AadTokenProvider is replacing the /_api.SP.OAuth.Token/Acquire endpoint
 * for authentication with ADAL.js. At some point in the near future, when Azure AD v2.0
 * can support the same scenarios as the original version, we will switch to MSAL.
 *
 * @public
 * @sealed
 */
export declare class AadTokenProvider implements IAadTokenProvider {
    /* Excluded from this release type: _tokenAcquisitionEventId */
    /* Excluded from this release type: _authContextManager */
    private _tokenAcquisitionEvent;
    private _fetchQueue;
    private _oboConfiguration;
    private readonly _defaultConfiguration;
    /* Excluded from this release type: __constructor */
    /**
     * Fetches the AAD OAuth2 token for a resource if the user that's currently logged in has
     * access to that resource.
     *
     * The OAuth2 token should not be cached by the caller since it is already cached by the method
     * itself.
     *
     * @param resourceEndpoint - the resource for which the token should be obtained
     * An example of a resourceEndpoint would be https://graph.microsoft.com
     * @returns A promise that will be fullfiled with the token or that will reject
     *          with an error message
     */
    getToken(resourceEndpoint: string): Promise<string>;
    /* Excluded from this release type: _getTokenInternal */
    /**
     * Notifies the developer when Token Acquisition requires user action.
     * @eventproperty
     */
    readonly tokenAcquisitionEvent: SPEvent<TokenAcquisitionEventArgs>;
}

/**
 * Returns a preinitialized version of the AadTokenProviderFactory.
 * @public
 */
export declare class AadTokenProviderFactory {
    /**
     * The service key for AadTokenProviderFactory.
     */
    static readonly serviceKey: ServiceKey<AadTokenProviderFactory>;
    /**
     * Returns an instance of the AadTokenProvider that communicates with the current tenant's configurable
     * Service Principal.
     */
    getTokenProvider(): Promise<AadTokenProvider>;
}

/* Excluded from this release type: _AadTokenProviders */

/* Excluded from this release type: AdalAuthContextManager */

/**
 * {@inheritDoc IDigestCache}
 *
 * @public
 */
export declare class DigestCache implements IDigestCache {
    /**
     * The service key for IDigestCache.
     */
    static readonly serviceKey: ServiceKey<IDigestCache>;
    private static REST_EXPIRATION_SLOP_MS;
    private static _logSource;
    private _fetchProvider;
    private _timeProvider;
    private _digestsByUrl;
    constructor(serviceScope: ServiceScope);
    /**
     * {@inheritDoc IDigestCache.fetchDigest}
     */
    fetchDigest(webUrl: string): Promise<string>;
    /**
     * {@inheritDoc IDigestCache.addDigestToCache}
     */
    addDigestToCache(webUrl: string, digestValue: string, expirationTimestamp: number): void;
    /**
     * {@inheritDoc IDigestCache.clearDigest}
     */
    clearDigest(webUrl: string): boolean;
    /**
     * {@inheritDoc IDigestCache.clearAllDigests}
     */
    clearAllDigests(): void;
}

/**
 * GraphHttpClient is used to perform REST calls against Microsoft Graph. It adds default
 * headers and collects telemetry that helps the service to monitor the performance of an application.
 * https://developer.microsoft.com/en-us/graph/
 *
 * @remarks
 * For communicating with SharePoint, use the {@link SPHttpClient} class instead.
 * For communicating with other internet services, use the {@link HttpClient} class instead.
 *
 * @beta
 * @deprecated The GraphHttpClient class has been superseded by the MSGraphClient class.
 * @sealed
 */
export declare class GraphHttpClient {
    /**
     * The standard predefined GraphHttpClientConfiguration objects for use with
     * the GraphHttpClient class.
     */
    static readonly configurations: IGraphHttpClientConfigurations;
    /**
     * The service key for GraphHttpClient.
     */
    static readonly serviceKey: ServiceKey<GraphHttpClient>;
    private static readonly _className;
    private static _logSource;
    private _graphContext;
    private _parentSource;
    private _serviceScope;
    private _token;
    private _tokenProvider;
    private _fetchProvider;
    constructor(serviceScope: ServiceScope);
    /**
     * Perform a REST service call.
     *
     * @remarks
     * Generally, the parameters and semantics for HttpClient.fetch() are essentially
     * the same as the WHATWG API standard that is documented here:
     * https://fetch.spec.whatwg.org/
     *
     * The GraphHttpClient subclass adds some additional behaviors that are convenient when
     * working with SharePoint ODATA API's (which can be avoided by using
     * HttpClient instead):
     * - Default "Accept" and "Content-Type" headers are added if not explicitly specified.
     *
     * @param url - The url string should be relative to the graph server.
     *  Good: 'v1.0/me/events'
     *  Bad: '/v1.0/me/events', 'https://graph.microsoft.com/v1.0/me/events'
     * @param configuration - determines the default behavior of GraphHttpClient; normally this should
     *   be the latest version number from GraphHttpClientConfigurations
     * @param options - additional options that affect the request
     * @returns a promise that will return the result
     */
    fetch(url: string, configuration: GraphHttpClientConfiguration, options: IGraphHttpClientOptions): Promise<GraphHttpClientResponse>;
    /**
     * Calls fetch(), but sets the method to "GET".
     * @param url - the URL to fetch
     * @param configuration - determines the default behavior of GraphHttpClient; normally this should
     *   be the latest version number from GraphHttpClientConfigurations
     * @param options - additional options that affect the request
     * @returns a promise that will return the result
     */
    get(url: string, configuration: GraphHttpClientConfiguration, options?: IGraphHttpClientOptions): Promise<GraphHttpClientResponse>;
    /**
     * Calls fetch(), but sets the method to "POST".
     * @param url - the URL to fetch
     * @param configuration - determines the default behavior of GraphHttpClient; normally this should
     *   be the latest version number from GraphHttpClientConfigurations
     * @param options - additional options that affect the request
     * @returns a promise that will return the result
     */
    post(url: string, configuration: GraphHttpClientConfiguration, options: IGraphHttpClientOptions): Promise<GraphHttpClientResponse>;
    private _fetchWithInstrumentation;
    private _getOAuthToken;
    private readonly _logSourceId;
    private _mergeUserHeaders;
    private _writeQosMonitorUpdate;
    /**
     * This function verifies that
     */
    private _validateGraphRelativeUrl;
}

/**
 * Configuration for {@link GraphHttpClient}.
 *
 * @remarks
 * The GraphHttpClientConfiguration object provides a set of switches for enabling/disabling
 * various features of the GraphHttpClient class.  Normally these switches are set
 * (e.g. when calling GraphHttpClient.fetch()) by providing one of the predefined defaults
 * from GraphHttpClientConfigurations, however switches can also be changed via the
 * GraphHttpClientConfiguration.overrideWith() method.
 *
 * @deprecated The GraphHttpClient class has been superceded by the MSGraphClient class.
 * @beta
 */
export declare class GraphHttpClientConfiguration extends HttpClientConfiguration implements IGraphHttpClientConfiguration {
    protected flags: IGraphHttpClientConfiguration;
    /**
     * Constructs a new instance of GraphHttpClientCommonConfiguration with the specified flags.
     *
     * @remarks
     * The default values will be used for any flags that are missing or undefined.
     * If overrideFlags is specified, it takes precedence over flags.
     */
    constructor(flags: IGraphHttpClientConfiguration, overrideFlags?: IGraphHttpClientConfiguration);
}

/* Excluded from this release type: _GraphHttpClientContext */

/**
 * The Response subclass returned by methods such as GraphHttpClient.fetch().
 *
 * @remarks
 * This is a placeholder.  In the future, additional GraphHttpClient-specific functionality
 * may be added to this class.
 *
 * @deprecated The GraphHttpClient class has been superceded by the MSGraphClient class.
 * @beta
 * @sealed
 */
export declare class GraphHttpClientResponse extends HttpClientResponse {
    private _correlationId;
    constructor(response: Response);
    /**
     * @override
     */
    clone(): GraphHttpClientResponse;
    /**
     * Returns the Graph API correlation ID.
     *
     * @returns the correlation ID, or undefined if the "request-id" header was not found.
     */
    readonly correlationId: Guid | undefined;
}

/**
 * Typings for the GraphRequest Object
 * For more information: {@link https://github.com/microsoftgraph/msgraph-sdk-javascript}
 *
 * @privateRemarks
 * The original typings include references to polyfills.
 *
 * @public
 */
export declare class GraphRequest {
    config: Options;
    urlComponents: URLComponents;
    _headers: {
        [key: string]: string | number;
    };
    _responseType: string;
    constructor(config: Options, path: string);
    header(headerKey: string, headerValue: string): this;
    headers(headers: {
        [key: string]: string | number;
    }): this;
    parsePath(rawPath: string): void;
    private urlJoin;
    buildFullUrl(): string;
    version(v: string): GraphRequest;
    select(properties: string | string[]): GraphRequest;
    expand(properties: string | string[]): GraphRequest;
    orderby(properties: string | string[]): GraphRequest;
    filter(filterStr: string): GraphRequest;
    top(n: number): GraphRequest;
    skip(n: number): GraphRequest;
    skipToken(token: string): GraphRequest;
    count(count: boolean): GraphRequest;
    responseType(responseType: string): GraphRequest;
    private addCsvQueryParamater;
    delete(callback?: GraphRequestCallback): Promise<any>;
    patch(content: any, callback?: GraphRequestCallback): Promise<any>;
    post(content: any, callback?: GraphRequestCallback): Promise<any>;
    put(content: any, callback?: GraphRequestCallback): Promise<any>;
    create(content: any, callback?: GraphRequestCallback): Promise<any>;
    update(content: any, callback?: GraphRequestCallback): Promise<any>;
    del(callback?: GraphRequestCallback): Promise<any>;
    get(callback?: GraphRequestCallback): Promise<any>;
    private routeResponseToPromise;
    private handleFetch;
    private routeResponseToCallback;
    private sendRequestAndRouteResponse;
    getStream(callback: GraphRequestCallback): void;
    putStream(stream: any, callback: GraphRequestCallback): void;
    private getDefaultRequestHeaders;
    private configureRequest;
    query(queryDictionaryOrString: string | {
        [key: string]: string | number;
    }): GraphRequest;
    private createQueryString;
    private convertResponseType;
}

/**
 * HttpClient implements a basic set of features for performing REST operations against
 * a generic service.
 *
 * @remarks
 * For communicating with SharePoint, use the {@link SPHttpClient} class instead.
 *
 * @public
 */
export declare class HttpClient {
    /**
     * The standard predefined HttpClientConfiguration objects for use with
     * the HttpClient class.
     */
    static readonly configurations: IHttpClientConfigurations;
    /**
     * The service key for HttpClient.
     *
     * @public
     */
    static readonly serviceKey: ServiceKey<HttpClient>;
    private static readonly _className;
    /* Excluded from this release type: _serviceScope */
    private _fetchProvider;
    constructor(serviceScope: ServiceScope);
    /**
     * Performs a REST service call.
     *
     * @remarks
     * Although the SPHttpClient subclass adds additional enhancements, the parameters and semantics
     * for HttpClient.fetch() are essentially the same as the WHATWG API standard that is documented here:
     * https://fetch.spec.whatwg.org/
     *
     * @param url - the URL to fetch
     * @param configuration - determines the default behavior of HttpClient; normally this should
     *   be the latest version number from HttpClientConfigurations
     * @param options - additional options that affect the request
     * @returns A promise with behavior similar to WHATWG fetch().  This promise will resolve normally
     * (with {@link HttpClientResponse.ok} being false) for error status codes such as HTTP 404
     * or 500.  The promise will only reject for network failures or other errors that prevent communication
     * with the server.
     */
    fetch(url: string, configuration: HttpClientConfiguration, options: IHttpClientOptions): Promise<HttpClientResponse>;
    /**
     * Calls fetch(), but sets the method to "GET".
     *
     * @param url - the URL to fetch
     * @param configuration - determines the default behavior of HttpClient; normally this should
     *   be the latest version number from HttpClientConfigurations
     * @param options - additional options that affect the request
     * @returns A promise with behavior similar to WHATWG fetch().  This promise will resolve normally
     * (with {@link HttpClientResponse.ok} being false) for error status codes such as HTTP 404
     * or 500.  The promise will only reject for network failures or other errors that prevent communication
     * with the server.
     */
    get(url: string, configuration: HttpClientConfiguration, options?: IHttpClientOptions): Promise<HttpClientResponse>;
    /**
     * Calls fetch(), but sets the method to "POST".
     *
     * @param url - the URL to fetch
     * @param configuration - determines the default behavior of HttpClient; normally this should
     *   be the latest version number from HttpClientConfigurations
     * @param options - additional options that affect the request
     * @returns A promise with behavior similar to WHATWG fetch().  This promise will resolve normally
     * (with {@link HttpClientResponse.ok} being false) for error status codes such as HTTP 404
     * or 500.  The promise will only reject for network failures or other errors that prevent communication
     * with the server.
     */
    post(url: string, configuration: HttpClientConfiguration, options: IHttpClientOptions): Promise<HttpClientResponse>;
}

/**
 * Configuration for HttpClient.
 *
 * @remarks
 * The HttpClientConfiguration object provides a set of switches for enabling/disabling
 * various features of the HttpClient class.  Normally these switches are set
 * (e.g. when calling HttpClient.fetch()) by providing one of the predefined defaults
 * from HttpClientConfigurations, however switches can also be changed via the
 * HttpClientConfiguration.overrideWith() method.
 *
 * @public
 */
export declare class HttpClientConfiguration implements IHttpClientConfiguration {
    protected flags: IHttpClientConfiguration;
    /**
     * Constructs a new instance of HttpClientConfiguration with the specified flags.
     * The default values will be used for any flags that are missing or undefined.
     * If overrideFlags is specified, it takes precedence over flags.
     */
    constructor(flags: IHttpClientConfiguration, overrideFlags?: IHttpClientConfiguration);
    /**
     * Child classes should override this method to construct the child class type,
     * rather than the base class type.
     * @virtual
     */
    overrideWith(sourceFlags: IHttpClientConfiguration): HttpClientConfiguration;
    /**
     * Child classes should override this method to initialize the flags object.
     * @virtual
     */
    protected initializeFlags(): void;
    private _mergeFlags;
}

/**
 * The Response subclass returned by methods such as HttpClient.fetch().
 *
 * @remarks
 * This is a placeholder.  In the future, additional HttpClient-specific functionality
 * may be added to this class.
 *
 * @privateRemarks
 * This class exposes the same members as our typings for the browser's native
 * Response and Body classes, which is why we can say that it "implements" them.
 * It cannot actually inherit from Response because that class does not have a copy
 * constructor (because it would probably be inefficient to copy the response stream).
 *
 * @public
 */
export declare class HttpClientResponse implements Response, Body {
    protected nativeResponse: Response;
    /* Excluded from this release type: __constructor */
    /* Excluded from this release type: body */
    /**
     * See documentation for the browser's
     * {@link https://developer.mozilla.org/en-US/docs/Web/API/Body | Body} API
     */
    readonly bodyUsed: boolean;
    /**
     * See documentation for the browser's
     * {@link https://developer.mozilla.org/en-US/docs/Web/API/Body | Body} API
     */
    arrayBuffer(): Promise<ArrayBuffer>;
    /**
     * See documentation for the browser's
     * {@link https://developer.mozilla.org/en-US/docs/Web/API/Body | Body} API
     */
    blob(): Promise<Blob>;
    /**
     * See documentation for the browser's
     * {@link https://developer.mozilla.org/en-US/docs/Web/API/Body | Body} API
     */
    formData(): Promise<FormData>;
    /**
     * See documentation for the browser's
     * {@link https://developer.mozilla.org/en-US/docs/Web/API/Body | Body} API
     */
    json(): Promise<any>;
    /**
     * See documentation for the browser's
     * {@link https://developer.mozilla.org/en-US/docs/Web/API/Body | Body} API
     */
    text(): Promise<string>;
    /**
     * See documentation for the browser's
     * {@link https://developer.mozilla.org/en-US/docs/Web/API/Response | Response} API
     */
    readonly type: ResponseType;
    /* Excluded from this release type: redirected */
    /**
     * See documentation for the browser's
     * {@link https://developer.mozilla.org/en-US/docs/Web/API/Response | Response} API
     */
    readonly url: string;
    /**
     * See documentation for the browser's
     * {@link https://developer.mozilla.org/en-US/docs/Web/API/Response | Response} API
     */
    readonly status: number;
    /**
     * See documentation for the browser's
     * {@link https://developer.mozilla.org/en-US/docs/Web/API/Response | Response} API
     */
    readonly ok: boolean;
    /**
     * See documentation for the browser's
     * {@link https://developer.mozilla.org/en-US/docs/Web/API/Response | Response} API
     */
    readonly statusText: string;
    /**
     * See documentation for the browser's
     * {@link https://developer.mozilla.org/en-US/docs/Web/API/Response | Response} API
     */
    readonly headers: Headers;
    /**
     * @virtual
     */
    clone(): HttpClientResponse;
}

/**
 * Flags interface for HttpClientConfiguration.
 *
 * @public
 */
export declare interface IAadHttpClientConfiguration extends IHttpClientConfiguration {
}

/**
 * Standard configurations for AadHttpClient.
 *
 * @remarks
 * This interface provides standard predefined AadHttpClientConfiguration objects for use with
 * the AadHttpClient class.  In general, clients should choose the latest available
 * version number, which enables all the switches that are recommended for typical
 * scenarios.  (If new switches are introduced in the future, a new version number
 * will be introduced, which ensures that existing code will continue to function the
 * way it did at the time when it was tested.)
 *
 * @public
 */
export declare interface IAadHttpClientConfigurations {
    /**
     * This configuration turns off every feature switch for AadHttpClient. The fetch()
     * behavior will be essentially identical to the WHATWG standard API that
     * is documented here:
     * https://fetch.spec.whatwg.org/
     */
    readonly v1: AadHttpClientConfiguration;
}

/**
 * Interface for overriding the default behavior of AadHttpClient.
 *
 * @public
 */
export declare interface IAadHttpClientOptions {
    tokenProvider?: IAadTokenProvider;
    configuration?: IAadTokenProviderConfiguration;
}

/**
 * This class allows a developer to obtain OAuth2 tokens from Azure AD.
 *
 * OAuth2 tokens are used to authenticate the user from the SharePoint page
 * to other services such as PowerBI, Sway, Exchange, Yammer, etc.
 *
 * @remarks
 * AadTokenProvider is replacing the /_api.SP.OAuth.Token/Acquire endpoint
 * for authentication with ADAL.js. At some point in the near future, when Azure AD v2.0
 * can support the same scenarios as the original version, we will switch to MSAL.
 *
 * @public
 */
export declare interface IAadTokenProvider {
    /**
     * Notifies the developer when Token Acquistion requires user action.
     * @eventproperty
     */
    readonly tokenAcquisitionEvent: SPEvent<ITokenAcquisitionEventArgs>;
    /**
     * Fetches the AAD OAuth2 token for a resource if the user that's currently logged in has
     * access to that resource.
     *
     * The OAuth2 token SHOULD NOT be cached by the caller since it is already cached by the method
     * itself.
     *
     * An example of a resourceEndpoint would be https://sdfpilot.outlook.com
     *
     * @param resourceEndpoint - the resource for which the token should be obtained
     * @returns A promise that will be fullfiled with the token or that will reject
     *          with an error message
     */
    getToken(resourceEndpoint: string): Promise<string>;
}

/**
 * Required strings for constructing an AadTokenProvider.
 *
 * @public
 */
export declare interface IAadTokenProviderConfiguration {
    /**
     * The sign in page used to authenticate with Azure Active Directory. Trailing slashes are forbidden.
     */
    aadInstanceUrl: string;
    /**
     * The Azure Active Directory's tenant id.
     */
    aadTenantId: string;
    /**
     * The page used to retrieve tokens from Azure Active Directory. This url must be listed in
     * the developer's application redirect uris.
     */
    redirectUri: string;
    /**
     * The client ID of the developer's Azure Active Directory application.
     */
    servicePrincipalId: string;
    /**
     * The user's Azure Active Directory id. This will be used to ensure that a valid cached token is for
     * the current user.
     */
    aadUserId?: string;
    /**
     * The user's email address. This will be used to ensure that the current user's identity is used for
     * fetching auth tokens. This parameter will avoid the "Request is ambiguous: multiple user identities
     * are avaliable for the current request" error.
     */
    userEmail?: string;
}

/* Excluded from this release type: IAdalJsModule */

/**
 * IDigestCache is an internal service used by SPHttpClient to maintain a cache of request digests
 * for each SPWeb URL.  A request digest is a security token that the SharePoint server requires for
 * for any REST write operation, specified via the "X-RequestDigest" HTTP header.  It is obtained
 * by calling the "/_api/contextinfo" REST endpoint, and expires after a server configurable amount
 * of time.
 *
 * For more information, see the MSDN article
 * {@link https://msdn.microsoft.com/en-us/library/office/jj164022.aspx
 * | "Complete basic operations using SharePoint 2013 REST endpoints" }
 *
 * @public
 */
export declare interface IDigestCache {
    /**
     * Returns a digest string for the specified SPWeb URL.  If the cache already contains a usable value,
     * the promise is fulfilled immediately.  Otherwise, the promise will be pending and resolve after
     * an HTTP request obtains the digest, which will be added to the cache.
     * @param webUrl - The URL of the SPWeb that the API call will be issued to.
     * This may be a server-relative or absolute URL.
     * @returns       A promise that is fulfilled with the digest value.
     */
    fetchDigest(webUrl: string): Promise<string>;
    /**
     * Inserts a specific request digest value into the cache.  Normally this is unnecessary because
     * the framework will automatically issue a REST request to fetch the digest when necessary;
     * however, in advanced scenarios addDigestToCache() can be used to avoid the overhead of the
     * REST call.
     *
     * @param webUrl - The URL of the SPWeb that the API call will be issued to.
     * This may be a server-relative or absolute URL.
     * @param digestValue - The digest value, which is an opaque that must be generated  by the SharePoint server.
     * The syntax will look something like this: "0x0B85...2EAC,29 Jan 2016 01:23:45 -0000"
     * @param expirationTimestamp - A future point in time, as measured by performance.now(),
     * after which the digest value will no longer be valid.
     * NOTE: The expirationTime is a DOMHighResTimeStamp value whose units are
     * fractional milliseconds; for example, to specify an expiration
     * "5 seconds from right now", use performance.now()+5000.
     */
    addDigestToCache(webUrl: string, digestValue: string, expirationTimestamp: number): void;
    /**
     * Clears the cached digest for the specified SPWeb URL.  This operation is useful
     * e.g. if an error indicates that a digest was invalidated prior to its expiration time.
     *
     * @param webUrl - The URL of the SPWeb whose digest should be cleared. This may be
     * a server-relative or absolute URL.
     * @returns Returns true if a cache entry was found and deleted; false otherwise.
     */
    clearDigest(webUrl: string): boolean;
    /**
     * Clears all values from the cache.
     */
    clearAllDigests(): void;
}

/* Excluded from this release type: IFetchProvider */

/**
 * Flags interface for GraphHttpClientCommonConfiguration
 *
 * @deprecated The GraphHttpClient class has been superceded by the MSGraphClient class.
 * @beta
 */
export declare interface IGraphHttpClientConfiguration extends IHttpClientConfiguration {
}

/**
 * Standard configurations for GraphHttpClient.
 *
 * @remarks
 * This interface provides standard predefined GraphHttpClientConfiguration objects for use with
 * the GraphHttpClient class.  In general, clients should choose the latest available
 * version number, which enables all the switches that are recommended for typical
 * scenarios.  (If new switches are introduced in the future, a new version number
 * will be introduced, which ensures that existing code will continue to function the
 * way it did at the time when it was tested.)
 *
 * @deprecated The GraphHttpClient class has been superceded by the MSGraphClient class.
 * @beta
 */
export declare interface IGraphHttpClientConfigurations {
    /**
     * This configuration turns off every feature switch for HttpClient.  The fetch()
     * behavior will be essentially identical to the WHATWG standard API that
     * is documented here:
     * https://fetch.spec.whatwg.org/
     */
    readonly v1: GraphHttpClientConfiguration;
}

/**
 * Options for HttpClient
 *
 * @remarks
 * This interface defines the options for the GraphHttpClient operations such as
 * get(), post(), fetch(), etc.  It is based on the WHATWG API standard
 * parameters that are documented here:
 * https://fetch.spec.whatwg.org/
 *
 * @beta
 */
export declare interface IGraphHttpClientOptions extends IHttpClientOptions {
}

/**
 * Flags interface for HttpClientConfiguration.
 *
 * @public
 */
export declare interface IHttpClientConfiguration {
}

/**
 * Standard configurations for HttpClient.
 *
 * @remarks
 * This interface provides standard predefined HttpClientConfiguration objects for use with
 * the HttpClient class.  In general, clients should choose the latest available
 * version number, which enables all the switches that are recommended for typical
 * scenarios.  (If new switches are introduced in the future, a new version number
 * will be introduced, which ensures that existing code will continue to function the
 * way it did at the time when it was tested.)
 *
 * @public
 */
export declare interface IHttpClientConfigurations {
    /**
     * This configuration turns off every feature switch for HttpClient.  The fetch()
     * behavior will be essentially identical to the WHATWG standard API that
     * is documented here:
     * https://fetch.spec.whatwg.org/
     */
    readonly v1: HttpClientConfiguration;
}

/**
 * Options for HttpClient
 *
 * @remarks
 * This interface defines the options for the HttpClient operations such as
 * get(), post(), fetch(), etc.  It is based on the whatwg API standard
 * parameters that are documented here:
 * https://fetch.spec.whatwg.org/
 *
 * @public
 */
export declare interface IHttpClientOptions extends RequestInit {
}

/**
 * Flags interface for SPHttpClientBatchConfiguration.
 *
 * @beta
 */
export declare interface ISPHttpClientBatchConfiguration extends ISPHttpClientCommonConfiguration {
}

/**
 * Standard configurations for SPHttpClient.
 *
 * @remarks
 * This interface provides standard predefined SPHttpClientBatchConfiguration objects for use with
 * the SPHttpClientBatch class.  In general, clients should choose the latest available
 * version number, which enables all the switches that are recommended for typical
 * scenarios.  (If new switches are introduced in the future, a new version number
 * will be introduced, which ensures that existing code will continue to function the
 * way it did at the time when it was tested.)
 *
 * @beta
 */
export declare interface ISPHttpClientBatchConfigurations {
    /**
     * Version 1 enables these switches:
     * consoleLogging = true;
     * jsonRequest = true;
     * jsonResponse = true
     */
    readonly v1: SPHttpClientBatchConfiguration;
}

/**
 * This interface is passed to the SPHttpClientBatch constructor.  It specifies options
 * that affect the entire batch.
 *
 * @beta
 */
export declare interface ISPHttpClientBatchCreationOptions {
    /**
     * SPHttpClientBatch will need to perform its POST to an endpoint such as
     * "http://example.com/sites/sample/_api/$batch". Typically the SPWeb URL
     * ("https://example.com/sites/sample" in this example) can be guessed by
     * looking for a reserved URL segment such as "_api" in the first URL
     * passed to fetch(), but if not, the webUrl can be explicitly specified
     * using this option.
     */
    webUrl?: string;
}

/**
 * This interface defines the options for an individual REST request that
 * is part of an SPHttpClientBatch.  It is based on the WHATWG API standard
 * parameters that are documented here:
 * https://fetch.spec.whatwg.org/
 *
 * @beta
 */
export declare interface ISPHttpClientBatchOptions extends IHttpClientOptions {
}

/**
 * Flags interface for SPHttpClientCommonConfiguration
 *
 * @public
 */
export declare interface ISPHttpClientCommonConfiguration extends IHttpClientConfiguration {
    /**
     * Automatically configure the "Content-Type" header for a JSON payload.
     *
     * @remarks
     * When this switch is true:
     *
     * If the "Content-Type" header was not explicitly added for the request,
     * then SPHttpClient will add it if the request is a write operation (i.e.
     * an HTTP method other than "GET", "HEAD", or "OPTIONS").
     *
     * For OData 3.0, the value is 'application/json;odata=verbose;charset=utf-8'.
     *
     * For OData 4.0, the value is 'application/json;charset=utf-8'.
     */
    jsonRequest?: boolean;
    /**
     * Automatically configure the "Accept" header for a JSON payload.
     *
     * @remarks
     * When this switch is true:
     *
     * If the "Accept" header was not explicitly added for the request,
     * then SPHttpClient will add it.
     *
     * For OData 3.0, the value is 'application/json'.
     *
     * For OData 4.0, the value is 'application/json;odata.metadata=minimal'.
     */
    jsonResponse?: boolean;
}

/**
 * Flags interface for SPHttpClientConfiguration.
 *
 * @public
 */
export declare interface ISPHttpClientConfiguration extends ISPHttpClientCommonConfiguration {
    /**
     * Automatically configure the RequestInit.credentials.
     *
     * @remarks
     * When this switch is true:
     *
     * If RequestInit.credentials is not explicitly specified for the request,
     * then SPHttpClient will assign it to be "same-origin".  Without this switch,
     * different web browsers may apply different defaults.
     *
     * For more information, see the spec:
     * https://fetch.spec.whatwg.org/#cors-protocol-and-credentials
     */
    defaultSameOriginCredentials?: boolean;
    /**
     * Automatically configure the "OData-Version" header.
     *
     * @remarks
     * When this switch is specified (i.e. not undefined):
     * If the "OData-Version" header was not explicitly added for the request,
     * then SPHttpClient will add the header to specify the version indicated
     * by defaultODataVersion.
     *
     * NOTE: Without an 'OData-Version' header, the SharePoint server currently
     * defaults to Version 3.0 in most cases.  The recommended version is 4.0.
     */
    defaultODataVersion?: ODataVersion;
    /**
     * Automatically provide an "X-RequestDigest" header for authentication.
     *
     * @remarks
     * When this switch is true:
     *
     * If the "X-RequestDigest" header was not explicitly added for the request,
     * then SPHttpClient will add it if the request is a write operation (i.e.
     * an HTTP method other than "GET", "HEAD", or "OPTIONS").  The request digest
     * is managed by the DigestCache service.  In the case of a cache miss, an
     * additional network request may be performed.
     */
    requestDigest?: boolean;
}

/**
 * Standard configurations for SPHttpClient.
 *
 * @remarks
 * This interface provides standard predefined SPHttpClientConfiguration objects for use with
 * the SPHttpClient class.  In general, clients should choose the latest available
 * version number, which enables all the switches that are recommended for typical
 * scenarios.  (If new switches are introduced in the future, a new version number
 * will be introduced, which ensures that existing code will continue to function the
 * way it did at the time when it was tested.)
 *
 * @public
 */
export declare interface ISPHttpClientConfigurations {
    /**
     * Version 1 enables these switches:
     *
     * consoleLogging = true;
     * jsonRequest = true;
     * jsonResponse = true;
     * defaultSameOriginCredentials = true;
     * defaultODataVersion = ODataVersion.v4;
     * requestDigest = true
     */
    readonly v1: SPHttpClientConfiguration;
}

/**
 * This interface defines the options for the SPHttpClient operations such as
 * get(), post(), fetch(), etc.  It is based on the WHATWG API standard
 * parameters that are documented here:
 * https://fetch.spec.whatwg.org/
 *
 * @public
 */
export declare interface ISPHttpClientOptions extends IHttpClientOptions {
    /**
     * Configure the SPWeb URL for authentication.
     *
     * @remarks
     * For a write operation, SPHttpClient will automatically add the
     * "X-RequestDigest" header, which may need to be fetched using a seperate
     * request such as "https://example.com/sites/sample/_api/contextinfo".
     * Typically the SPWeb URL ("https://example.com/sites/sample" in this
     * example) can be guessed by looking for a reserved URL segment such
     * as "_api" in the original REST query, however certain REST endpoints
     * do not contain a reserved URL segment; in this case, the webUrl can
     * be explicitly specified using this option.
     */
    webUrl?: string;
}

/* Excluded from this release type: ISPOBOFlowParameters */

/**
 * Represents arguments used for raising a token acquisiton failure event.
 * @public
 */
declare interface ITokenAcquisitionEventArgs extends SPEventArgs {
    /**
     * The message returned from ADAL fails to retrieve a token from Azure AD.
     */
    message: string;
    /**
     * The url of the page for the end user to interact with Azure AD.
     */
    redirectUrl?: string;
}

/* Excluded from this release type: ITokenProvider */

/**
 * MSGraphClient is used to perform REST calls against Microsoft Graph.
 *
 * @remarks The Microsoft Graph JavaScript client library is a lightweight wrapper around the
 * Microsoft Graph API. This class allows developers to start making REST calls to MSGraph without
 * needing to initialize the the MSGraph client library. If a custom configuration is desired,
 * the MSGraphClient api function needs to be provided with that custom configuration for
 * every request.
 *
 * For more information: {@link https://github.com/microsoftgraph/msgraph-sdk-javascript}
 *
 * @public
 */
export declare class MSGraphClient {
    private static _instance;
    private static _window;
    private static _originalConfig;
    private static _graphBaseUrl;
    /* Excluded from this release type: __constructor */
    /**
     * All calls to Microsoft Graph are chained together starting with the api function.
     *
     * @remarks Path supports the following formats:
     * * me
     * * /me
     * * https://graph.microsoft.com/v1.0/me
     * * https://graph.microsoft.com/beta/me
     * * me/events?$filter=startswith(subject, 'ship')
     *
     * The authProvider and baseUrl option should not be used, as they have already been
     * provided by the framework. See the official documentation here:
     * https://github.com/microsoftgraph/msgraph-sdk-javascript
     *
     * @param path - The path for the request to MSGraph.
     * @param config - Sets the configuration for this request.
     */
    api(path: string, config?: Options): GraphRequest;
    private _createGraphClientInstance;
    private _getOAuthToken;
}

/**
 * Returns a preinitialized version of the MSGraphClient.
 * For more information: {@link https://docs.microsoft.com/en-us/sharepoint/dev/spfx/use-msgraph}
 *
 * @public
 */
export declare class MSGraphClientFactory {
    /**
     * The service key for MSGraphClientFactory.
     */
    static readonly serviceKey: ServiceKey<MSGraphClientFactory>;
    private _serviceScope;
    /* Excluded from this release type: __constructor */
    /**
     * Returns an instance of the MSGraphClient that communicates with the current tenant's configurable
     * Service Principal.
     */
    getClient(): Promise<MSGraphClient>;
}

/* Excluded from this release type: OAuthTokenProvider */

/**
 * Represents supported version of the "OData-Version" header, which is part
 * of the Open Data Protocol standard.
 *
 * @public
 */
export declare class ODataVersion {
    /**
     * Represents version 3.0 for the "OData-Version" header
     */
    static v3: ODataVersion;
    /**
     * Represents version 4.0 for the "OData-Version" header
     */
    static v4: ODataVersion;
    private _versionString;
    /**
     * Attempt to parse the "OData-Version" header.
     *
     * @remarks
     * If the "OData-Version" header is present, this returns the
     * corresponding ODataVersion constant.  An error is thrown if
     * the version number is not supported.  If the header is missing,
     * then undefined is returned.
     */
    static tryParseFromHeaders(headers: Headers): ODataVersion | undefined;
    /**
     * Returns the "OData-Version" value, for example "4.0".
     */
    toString(): string;
    private constructor();
}

/**
 * SPHttpClient is used to perform REST calls against SharePoint.  It adds default
 * headers, manages the digest needed for writes, and collects telemetry that
 * helps the service to monitor the performance of an application.
 *
 * @remarks
 * For communicating with other internet services, use the {@link HttpClient} class.
 *
 * @public
 * @sealed
 */
export declare class SPHttpClient {
    /**
     * The standard predefined SPHttpClientConfiguration objects for use with
     * the SPHttpClient class.
     */
    static readonly configurations: ISPHttpClientConfigurations;
    /**
     * The service key for SPHttpClient.
     */
    static readonly serviceKey: ServiceKey<SPHttpClient>;
    private static readonly _className;
    private static _logSource;
    private static _reservedUrlSegments;
    private _digestCache;
    private _parentSource;
    private _serviceScope;
    private _fetchProvider;
    /**
     * Use a heuristic to infer the base URL for authentication.
     *
     * @remarks
     * Attempts to infer the SPWeb URL associated with the provided REST URL, by looking
     * for common SharePoint path components such as "_api", "_layouts", or "_vit_bin".
     * This is necessary for operations such as the X-RequestDigest
     * and ODATA batching, which require POSTing to a separate REST endpoint
     * in order to complete a request.
     *
     * For example, if the requestUrl is "/sites/site/web/_api/service",
     * the returned URL would be "/sites/site/web".  Or if the requestUrl
     * is "http://example.com/_layouts/service", the returned URL would be
     * "http://example.com".
     *
     * If the URL cannot be determined, an exception is thrown.
     *
     * @param requestUrl - The URL for a SharePoint REST service
     * @returns the inferred SPWeb URL
     */
    static getWebUrlFromRequestUrl(requestUrl: string): string;
    constructor(serviceScope: ServiceScope);
    /**
     * Perform a REST service call.
     *
     * @remarks
     * Generally, the parameters and semantics for SPHttpClient.fetch() are essentially
     * the same as the WHATWG API standard that is documented here:
     * https://fetch.spec.whatwg.org/
     *
     * The SPHttpClient subclass adds some additional behaviors that are convenient when
     * working with SharePoint ODATA API's (which can be avoided by using
     * HttpClient instead):
     *
     * - Default "Accept" and "Content-Type" headers are added if not explicitly specified.
     *
     * - For write operations, an "X-RequestDigest" header is automatically added
     *
     * - The request digest token is automatically fetched and stored in a cache, with
     *   support for preloading
     *
     * For a write operation, SPHttpClient will automatically add the "X-RequestDigest"
     * header, which may need to be obtained by issuing a seperate request such as
     * "https://example.com/sites/sample/_api/contextinfo".  Typically the appropriate
     * SPWeb URL can be guessed by looking for a reserved URL segment such as "_api"
     * in the original URL passed to fetch(); if not, use ISPHttpClientOptions.webUrl
     * to specify it explicitly.
     *
     * @param url - the URL to fetch
     * @param configuration - determines the default behavior of SPHttpClient; normally this should
     *   be the latest version number from SPHttpClientConfigurations
     * @param options - additional options that affect the request
     * @returns A promise with behavior similar to WHATWG fetch().  This promise will resolve normally
     * (with {@link HttpClientResponse.ok} being false) for error status codes such as HTTP 404
     * or 500.  The promise will only reject for network failures or other errors that prevent communication
     * with the server.
     */
    fetch(url: string, configuration: SPHttpClientConfiguration, options: ISPHttpClientOptions): Promise<SPHttpClientResponse>;
    /**
     * Calls fetch(), but sets the method to "GET".
     *
     * @param url - the URL to fetch
     * @param configuration - determines the default behavior of SPHttpClient; normally this should
     *   be the latest version number from SPHttpClientConfigurations
     * @param options - additional options that affect the request
     * @returns A promise with behavior similar to WHATWG fetch().  This promise will resolve normally
     * (with {@link HttpClientResponse.ok} being false) for error status codes such as HTTP 404
     * or 500.  The promise will only reject for network failures or other errors that prevent communication
     * with the server.
     */
    get(url: string, configuration: SPHttpClientConfiguration, options?: ISPHttpClientOptions): Promise<SPHttpClientResponse>;
    /**
     * Calls fetch(), but sets the method to "POST".
     *
     * @param url - the URL to fetch
     * @param configuration - determines the default behavior of SPHttpClient; normally this should
     *   be the latest version number from SPHttpClientConfigurations
     * @param options - additional options that affect the request
     * @returns A promise with behavior similar to WHATWG fetch().  This promise will resolve normally
     * (with {@link HttpClientResponse.ok} being false) for error status codes such as HTTP 404
     * or 500.  The promise will only reject for network failures or other errors that prevent communication
     * with the server.
     */
    post(url: string, configuration: SPHttpClientConfiguration, options: ISPHttpClientOptions): Promise<SPHttpClientResponse>;
    /**
     * Begins an ODATA batch, which allows multiple REST queries to be bundled into
     * a single web request.
     *
     * @returns An {@link SPHttpClientBatch} object used to manage the batch operation.
     *
     * @beta
     */
    beginBatch(batchCreationOptions?: ISPHttpClientBatchCreationOptions): SPHttpClientBatch;
    private _fetchWithInstrumentation;
    private readonly _logSourceId;
}

/**
 * The SPHttpClientBatch class accumulates a number of REST service calls and
 * transmits them as a single ODATA batch.  This protocol is documented here:
 * http://docs.oasis-open.org/odata/odata/v4.0/odata-v4.0-part1-protocol.html
 *
 * The usage is to call SPHttpClientBatch.fetch() to queue each individual request,
 * and then call SPHttpClientBatch.execute() to execute the batch operation.
 * The execute() method returns a promise that resolves when the real REST
 * call has completed.  Each call to fetch() also returns a promise that will
 * resolve with an SPHttpClientResponse object for that particular request.
 *
 * @privateRemarks
 * The type signature of SPHttpClientBatch class suggests that it should inherit from
 * the HttpClient base class.  However, the operational semantics are different
 * (e.g. nothing happens until execute() is called; further operations are
 * prohibited afterwards; fetch() calls cannot depend on each other).  In the
 * future we might introduce a base class for batches, but it would be separate
 * from the HttpClient hierarchy.  By contrast, the ISPHttpClientBatchOptions
 * does naturally inherit from IHttpClientOptions.
 *
 * @beta
 */
export declare class SPHttpClientBatch {
    /**
     * The standard predefined SPHttpClientBatchConfigurations objects for use with
     * the SPHttpClientBatch class.
     */
    static readonly configurations: ISPHttpClientBatchConfigurations;
    private _fetchProvider;
    private _randomNumberGenerator;
    private _digestCache;
    private _batchedRequests;
    private _webUrl;
    /* Excluded from this release type: __constructor */
    /**
     * Queues a new request, and returns a promise that can be used to access
     * the server response (after execute() has completed).
     *
     * @remarks
     * The parameters for this function are basically the same as the WHATWG API standard
     * documented here:
     *
     * {@link https://fetch.spec.whatwg.org/ }
     *
     * However, be aware that certain REST headers are ignored or not allowed inside
     * a batch.  See the ODATA documentation for details.
     *
     * When execute() is called, it will POST to a URL such as
     * "http://example.com/sites/sample/_api/$batch".  Typically SPHttpClientBatch can successfully
     * guess the appropriate SPWeb URL by looking for a reserved URL segment such as "_api"
     * in the first URL passed to fetch().  If not, use ISPHttpClientBatchCreationOptions.webUrl to specify it
     * explicitly.
     *
     * @param url - the URL to fetch
     * @param configuration - determines the default behavior of this request; normally this should
     *   be the latest version number from SPHttpClientBatchConfigurations
     * @param options - additional options that affect the request
     *
     * @returns A promise with behavior similar to WHATWG fetch().  This promise will resolve normally
     * (with {@link HttpClientResponse.ok} being false) for error status codes such as HTTP 404
     * or 500.  The promise will only reject for network failures or other errors that prevent communication
     * with the server.
     */
    fetch(url: string, configuration: SPHttpClientBatchConfiguration, options?: ISPHttpClientBatchOptions): Promise<SPHttpClientResponse>;
    /**
     * Calls fetch(), but sets the method to 'GET'.
     * @param url - the URL to fetch
     * @param configuration - determines the default behavior of this request; normally this should
     *   be the latest version number from SPHttpClientBatchConfigurations
     * @param options - additional options that affect the request
     * @returns A promise with behavior similar to WHATWG fetch().  This promise will resolve normally
     * (with {@link HttpClientResponse.ok} being false) for error status codes such as HTTP 404
     * or 500.  The promise will only reject for network failures or other errors that prevent communication
     * with the server.
     */
    get(url: string, configuration: SPHttpClientBatchConfiguration, options?: ISPHttpClientBatchOptions): Promise<SPHttpClientResponse>;
    /**
     * Calls fetch(), but sets the method to 'POST'.
     * @param url - the URL to fetch
     * @param configuration - determines the default behavior of this request; normally this should
     *   be the latest version number from SPHttpClientBatchConfigurations
     * @param options - additional options that affect the request
     * @returns A promise with behavior similar to WHATWG fetch().  This promise will resolve normally
     * (with {@link HttpClientResponse.ok} being false) for error status codes such as HTTP 404
     * or 500.  The promise will only reject for network failures or other errors that prevent communication
     * with the server.
     */
    post(url: string, configuration: SPHttpClientBatchConfiguration, options: ISPHttpClientBatchOptions): Promise<SPHttpClientResponse>;
    /**
     * Executes the batched queries that were queued using SPHttpClientBatch.fetch().
     */
    execute(): Promise<SPHttpClientBatch>;
    private _parseResponsesFromBody;
}

/**
 * Configuration for SPHttpClientBatch.
 *
 * @remarks
 * The SPHttpClientBatchConfiguration object provides a set of switches for enabling/disabling
 * various features of the SPHttpClientBatch class.  Normally these switches are set
 * (e.g. when calling SPHttpClientBatch.fetch()) by providing one of the predefined defaults
 * from SPHttpClientBatchConfigurations, however switches can also be changed via the
 * SPHttpClientBatchConfiguration.overrideWith() method.
 *
 * @beta
 */
export declare class SPHttpClientBatchConfiguration extends SPHttpClientCommonConfiguration implements ISPHttpClientBatchConfiguration {
    protected flags: ISPHttpClientBatchConfiguration;
    /**
     * Constructs a new instance of SPHttpClientBatchConfiguration with the specified flags.
     * The default values will be used for any flags that are missing or undefined.
     * If overrideFlags is specified, it takes precedence over flags.
     */
    constructor(flags: ISPHttpClientBatchConfiguration, overrideFlags?: ISPHttpClientBatchConfiguration);
    /**
     * @override
     */
    overrideWith(sourceFlags: ISPHttpClientBatchConfiguration): SPHttpClientBatchConfiguration;
    /**
     * @override
     */
    protected initializeFlags(): void;
}

/**
 * Common base class for SPHttpClientConfiguration and SPHttpClientBatchConfiguration.
 *
 * @public
 */
export declare class SPHttpClientCommonConfiguration extends HttpClientConfiguration implements ISPHttpClientCommonConfiguration {
    protected flags: ISPHttpClientCommonConfiguration;
    /**
     * Constructs a new instance of SPHttpClientCommonConfiguration with the specified flags.
     *
     * @remarks
     * The default values will be used for any flags that are missing or undefined.
     * If overrideFlags is specified, it takes precedence over flags.
     */
    constructor(flags: ISPHttpClientCommonConfiguration, overrideFlags?: ISPHttpClientCommonConfiguration);
    /**
     * @override
     */
    overrideWith(sourceFlags: ISPHttpClientCommonConfiguration): SPHttpClientCommonConfiguration;
    /**
     * {@inheritDoc ISPHttpClientCommonConfiguration.jsonRequest}
     */
    readonly jsonRequest: boolean;
    /**
     * {@inheritDoc ISPHttpClientCommonConfiguration.jsonResponse}
     */
    readonly jsonResponse: boolean;
    /**
     * @override
     */
    protected initializeFlags(): void;
}

/**
 * Configuration for {@link SPHttpClient}.
 *
 * @remarks
 * The SPHttpClientConfiguration object provides a set of switches for enabling/disabling
 * various features of the SPHttpClient class.  Normally these switches are set
 * (e.g. when calling SPHttpClient.fetch()) by providing one of the predefined defaults
 * from SPHttpClientConfigurations, however switches can also be changed via the
 * SPHttpClientConfiguration.overrideWith() method.
 *
 * @public
 */
export declare class SPHttpClientConfiguration extends SPHttpClientCommonConfiguration implements ISPHttpClientConfiguration {
    protected flags: ISPHttpClientConfiguration;
    /**
     * Constructs a new instance of SPHttpClientConfiguration with the specified flags.
     * The default values will be used for any flags that are missing or undefined.
     * If overrideFlags is specified, it takes precedence over flags.
     */
    constructor(flags: ISPHttpClientConfiguration, overrideFlags?: ISPHttpClientConfiguration);
    /**
     * @override
     */
    overrideWith(sourceFlags: ISPHttpClientConfiguration): SPHttpClientConfiguration;
    /**
     * {@inheritDoc ISPHttpClientConfiguration.defaultSameOriginCredentials}
     */
    readonly defaultSameOriginCredentials: boolean;
    /**
     * {@inheritDoc ISPHttpClientConfiguration.defaultODataVersion}
     */
    readonly defaultODataVersion: ODataVersion;
    /**
     * {@inheritDoc ISPHttpClientConfiguration.requestDigest}
     */
    readonly requestDigest: boolean;
    /**
     * @override
     */
    protected initializeFlags(): void;
}

/**
 * The Response subclass returned by methods such as SPHttpClient.fetch().
 *
 * @remarks
 * This is a placeholder.  In the future, additional SPHttpClient-specific functionality
 * may be added to this class.
 *
 * @public
 * @sealed
 */
export declare class SPHttpClientResponse extends HttpClientResponse {
    private _correlationId;
    constructor(response: Response);
    /**
     * @override
     */
    clone(): SPHttpClientResponse;
    /**
     * Returns the SharePoint correlation ID.
     *
     * @remarks
     *
     * The correlation ID is a Guid that can be used to associate log events that
     * are part of the same overall operation, but may originate from different services
     * or components.  SharePoint REST operations return the server's correlation ID
     * as the "sprequestguid" header.
     *
     * @returns the correlation ID, or undefined if the "sprequestguid" header was not found
     *
     * @beta
     */
    readonly correlationId: Guid | undefined;
    /* Excluded from this release type: statusMessage */
}

/**
 * Standard HTTP headers used with {@link SPHttpClient}
 *
 * @beta
 */
export declare const enum SPHttpHeader {
    /**
     * SharePoint uses the 'SPRequestGuid' header to return the server's correlation ID
     * for a request.
     *
     * Example value: "9417279e-40e1-0000-2465-306ba786bfd7"
     */
    SPRequestGuid = "SPRequestGuid"
}

/**
 * Arguments for a token acquisition failure event.
 * @public
 */
export declare class TokenAcquisitionEventArgs extends SPEventArgs {
    /**
     * The message returned from ADAL fails to retrieve a token from Azure AD.
     */
    message: string;
    /**
   * The url of the page for the end user to perform Multi Factor Authentication
   */
    redirectUrl?: string;
    constructor(message: string, redirectUrl?: string);
}

export { }
