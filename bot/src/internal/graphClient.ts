import { Client } from '@microsoft/microsoft-graph-client';
import { BearerTokenAuthProvider } from './bearerTokenAuthProvider';
import "isomorphic-fetch";

export class GraphClient {
    private _accessToken: string;
    private _authProvider: BearerTokenAuthProvider;
    public Client: Client;

    constructor(accessToken: string) {
        if (!accessToken || !accessToken.trim()) {
            throw new Error('SimpleGraphClient: Invalid token received.');
        }

        this._accessToken = accessToken;
        this._authProvider = new BearerTokenAuthProvider(this._accessToken);

        this.Client = Client.initWithMiddleware({
            authProvider: this._authProvider
        });
    }
}