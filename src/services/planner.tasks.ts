

/**
 * We are using OfficeHelpers library as it allows us to complete authentication
 * with relative ease and also provides other useful Utilities.
 *
 * Note: We have included a beta version of OfficeHelpers that has support for
 * MicrosoftTeams and the API signatures might change when OfficeHelpers for
 * Microsoft Teams releases.
 */
import { Authenticator, IToken, Utilities } from '@microsoft/office-js-helpers';
import { IPlannerTask } from '../core';

interface IPlannerService {
    authenticator: Authenticator;
    login: () => Promise<IToken>;
    get: () => Promise<IPlannerTask[]>;
    create: (task: IPlannerTask) => Promise<IPlannerTask>;
    delete: (task: IPlannerTask) => Promise<boolean>;
    markAsComplete: (task: IPlannerTask) => Promise<IPlannerTask>;
}

export class PlannerTasks implements IPlannerService {
    private _token: IToken;
    private _baseUrl: string = 'https://outlook.office.com/api/v2.0';
    private _login: boolean;

    authenticator: Authenticator;

    /**
     * Note: This is a demo clientID and can be removed/deactivated
     * at any point of time. Please modify this to use your own
     * clientId.
     */
    clientId = 'd560431b-2b07-4553-a24c-e0075fc3bbb6';

    constructor() {

        /**
         * OfficeHelpers provides a authentication utility that does Implicit OAuth
         * in simple steps and supports Microsoft, AzureAD, Google, Facebook out of the box.
         *
         * Note: We have included a beta version of OfficeHelpers that has support for
         * MicrosoftTeams and the API signatures might change when OfficeHelpers for
         * Microsoft Teams releases.
         *
         * If you wish to achieve the same authentication call without OfficeHelpers,
         * then you would need to call the following Microsoft Teams API manually and
         * handle token mangement yourself.
         *
         * microsoftTeams.authentication.authenticate({
         *     url: <constructed OAuth url to the authentication provider with required parameters>,
         *     width: <width in pixels>,
         *     height: <height in pixels>,
         *     successCallback: <handle message from dialog here>,
         *     failureCallback: <handle error from dialog here>
         * });
         *
         * Inside the dialog you can call
         * microsoftTeams.authentication.notifySuccess(<message string to be sent back to app>)
         * OR
         * microsoftTeams.authentication.notifyFailure(<error string to be sent back to app>)
         * accordingly.
         */
        this.authenticator = new Authenticator();
        // this.authenticator.endpoints.registerMicrosoftAuth(this.clientId, {
        //     baseUrl: 'https://login.microsoftonline.com/organizations/oauth2/v2.0',
        //     scope: 'https://outlook.office.com/tasks.readwrite'
        // });
        this.authenticator.endpoints.registerAzureADAuth(this.clientId, 'sanitariumdev.onmicrosoft.com');

        debugger;
        this._token = this.authenticator.tokens.get('AzureAD');
        if (this._token == null) {
            this.login();
        }
    }

    async login(): Promise<IToken> {
        /**
         * Read the comments above to achieve Authentication without depending on
         * the OfficeHelpers library.
         */
        try {
            debugger;
            this._token = await this.authenticator.authenticate('AzureAD');
            return this._token;
        }
        catch (error) {
            Utilities.log(error);
            throw new Error('Failed to login using your Work Account');
        };
    }

    logout() {
        this.authenticator.tokens.clear();
    }

    async get(): Promise<IPlannerTask[]> {
        try {
            let res = await fetch(`${this._baseUrl}/me/tasks`,
                {
                    headers: {
                        'Authorization': `Bearer ${this._token.access_token}`
                    }
                });

            res = await res.json();
            let tasks = this._checkForErrors(res);
            return tasks.value.map(this._convertTask);
        }
        catch (error) {
            Utilities.log(error);
            throw new Error('Failed to get your todos');
        };
    }

    async create(todo: IPlannerTask): Promise<IPlannerTask> {
        try {
            let res = await fetch(`${this._baseUrl}/me/tasks`,
                {
                    method: 'POST',
                    body: JSON.stringify({
                        Subject: todo.title,
                        Status: todo.completed ? 'Completed' : 'NotStarted',
                        Importance: todo.importance == null ? 'normal' : todo.importance
                    }),
                    headers: {
                        'Authorization': `Bearer ${this._token.access_token}`,
                        'Content-Type': 'application/json',
                        'Accept': 'application/json'
                    }
                });

            res = await res.json();
            let task = this._checkForErrors(res);
            return this._convertTask(task);
        }
        catch (error) {
            Utilities.log(error);
            throw new Error('Failed to create your todo');
        };
    }

    async update(todo: IPlannerTask): Promise<IPlannerTask> {
        try {
            let res = await fetch(`${this._baseUrl}/me/tasks('${todo.id}')`,
                {
                    method: 'PATCH',
                    body: JSON.stringify({
                        Subject: todo.title,
                        Status: todo.completed ? 'Completed' : 'NotStarted',
                        Importance: todo.importance == null ? 'normal' : todo.importance
                    }),
                    headers: {
                        'Authorization': `Bearer ${this._token.access_token}`,
                        'Content-Type': 'application/json',
                        'Accept': 'application/json'
                    }
                });

            res = await res.json();
            let task = this._checkForErrors(res);
            return this._convertTask(task);
        }
        catch (error) {
            Utilities.log(error);
            throw new Error('Failed to create your todo');
        };
    }

    async delete(todo: IPlannerTask): Promise<boolean> {
        try {
            let res = await fetch(`${this._baseUrl}/me/tasks('${todo.id}')`,
                {
                    method: 'DELETE',
                    headers: {
                        'Authorization': `Bearer ${this._token.access_token}`,
                        'Content-Type': 'application/json',
                        'Accept': 'application/json'
                    }
                });

            if (res.status === 204) {
                return this._checkForErrors(res);
            }
        }
        catch (error) {
            Utilities.log(error);
            throw new Error('Failed to delete your todo');
        };
    }

    async markAsComplete(task: IPlannerTask): Promise<IPlannerTask> {
        try {
            let res = await fetch(`${this._baseUrl}/me/tasks('${task.id}')/complete`,
                {
                    method: 'POST',
                    headers: {
                        'Authorization': `Bearer ${this._token.access_token}`,
                        'Content-Type': 'application/json',
                        'Accept': 'application/json'
                    }
                });

            res = await res.json();
            let task = this._checkForErrors(res);
            return this._convertTask(task);
        }
        catch (error) {
            Utilities.log(error);
            throw new Error('Failed to mark your todo as complete');
        }
    }

    private _convertTask(task: any): IPlannerTask {
        return {
            id: task.Id,
            title: task.Subject,
            completed: task.Status === 'Completed',
            importance: task.Importance && task.Importance.toLowerCase()
        } as IPlannerTask;
    }

    private _checkForErrors = response => {
        if (response.error) {
            throw new Error(response.error.message)
        }

        return response;
    }
}