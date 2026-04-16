// <copyright file="httpClient.ts" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

import * as microsoftTeams from "@microsoft/teams-js";
import i18n from '../i18n';

/**
 * Response shape that mirrors AxiosResponse so all existing consumers
 * (response.data, response.status, response.headers) keep working.
 */
export interface ApiResponse<T = any> {
    data: T;
    status: number;
    statusText: string;
    headers: Record<string, string>;
}

/**
 * Error shape that mirrors axios error so catch blocks using
 * error.response.status / error.response.data keep working.
 */
export class ApiError extends Error {
    public response: ApiResponse;
    constructor(response: ApiResponse) {
        super(`Request failed with status ${response.status}`);
        this.name = "ApiError";
        this.response = response;
    }
}

/** Headers dictionary used internally. */
interface HeadersMap {
    [key: string]: string;
}

export class AxiosJWTDecorator {
    public async get<T = any>(
        url: string,
        handleError: boolean = true,
        needAuthorizationHeader: boolean = true,
    ): Promise<ApiResponse<T>> {
        try {
            const headers = needAuthorizationHeader
                ? await this.getAuthHeaders()
                : this.getDefaultHeaders();
            return await this.request<T>(url, { method: "GET", headers });
        } catch (error) {
            if (handleError) { this.handleError(error); }
            throw error;
        }
    }

    public async patch<T = any>(
        url: string,
        data?: any,
        handleError: boolean = true,
    ): Promise<ApiResponse<T>> {
        try {
            const headers = await this.getAuthHeaders();
            return await this.request<T>(url, {
                method: "PATCH",
                headers: { ...headers, "Content-Type": "application/json" },
                body: data != null ? JSON.stringify(data) : undefined,
            });
        } catch (error) {
            if (handleError) { this.handleError(error); }
            throw error;
        }
    }

    public async delete<T = any>(
        url: string,
        handleError: boolean = true,
    ): Promise<ApiResponse<T>> {
        try {
            const headers = await this.getAuthHeaders();
            return await this.request<T>(url, { method: "DELETE", headers });
        } catch (error) {
            if (handleError) { this.handleError(error); }
            throw error;
        }
    }

    public async post<T = any>(
        url: string,
        data?: any,
        handleError: boolean = true,
    ): Promise<ApiResponse<T>> {
        try {
            const headers = await this.getAuthHeaders();
            return await this.request<T>(url, {
                method: "POST",
                headers: { ...headers, "Content-Type": "application/json" },
                body: data != null ? JSON.stringify(data) : undefined,
            });
        } catch (error) {
            if (handleError) { this.handleError(error); }
            throw error;
        }
    }

    public async put<T = any>(
        url: string,
        data?: any,
        handleError: boolean = true,
    ): Promise<ApiResponse<T>> {
        try {
            const headers = await this.getAuthHeaders();
            return await this.request<T>(url, {
                method: "PUT",
                headers: { ...headers, "Content-Type": "application/json" },
                body: data != null ? JSON.stringify(data) : undefined,
            });
        } catch (error) {
            if (handleError) { this.handleError(error); }
            throw error;
        }
    }

    // ── Private helpers ──────────────────────────────────────────────

    /**
     * Core fetch wrapper. Returns an ApiResponse (same shape as AxiosResponse).
     * Throws ApiError for non-2xx responses (same shape as axios error).
     */
    private async request<T>(url: string, init: RequestInit): Promise<ApiResponse<T>> {
        const response = await fetch(url, init);

        // Parse headers into a plain object.
        const headers: Record<string, string> = {};
        response.headers.forEach((value, key) => { headers[key] = value; });

        // Try to parse body as JSON; fall back to text then null.
        let data: any;
        const contentType = response.headers.get("content-type") || "";
        if (contentType.includes("application/json")) {
            data = await response.json();
        } else {
            const text = await response.text();
            data = text || null;
        }

        const apiResponse: ApiResponse<T> = {
            data,
            status: response.status,
            statusText: response.statusText,
            headers,
        };

        if (!response.ok) {
            throw new ApiError(apiResponse);
        }

        return apiResponse;
    }

    private handleError(error: any): void {
        if (error && error.response) {
            const errorStatus = error.response.status;
            if (errorStatus === 403) {
                window.location.href = `/errorpage/403?locale=${i18n.language}`;
            } else if (errorStatus === 401) {
                window.location.href = `/errorpage/401?locale=${i18n.language}`;
            } else {
                window.location.href = `/errorpage?locale=${i18n.language}`;
            }
        } else {
            window.location.href = `/errorpage?locale=${i18n.language}`;
        }
    }

    private getDefaultHeaders(): HeadersMap {
        return {
            "Accept-Language": i18n.language,
        };
    }

    private async getAuthHeaders(): Promise<HeadersMap> {
        microsoftTeams.initialize();

        return new Promise<HeadersMap>((resolve, reject) => {
            const authTokenRequest = {
                successCallback: (token: string) => {
                    resolve({
                        "Authorization": `Bearer ${token}`,
                        "Accept-Language": i18n.language,
                    });
                },
                failureCallback: (error: string) => {
                    // When the getAuthToken function returns a "resourceRequiresConsent" error,
                    // it means Azure AD needs the user's consent before issuing a token to the app.
                    // The following code redirects the user to the "Sign in" page where the user can grant the consent.
                    // Right now, the app redirects to the consent page for any error.
                    console.error("Error from getAuthToken: ", error);
                    window.location.href = `/signin?locale=${i18n.language}`;
                },
                resources: [] as any[]
            };
            microsoftTeams.authentication.getAuthToken(authTokenRequest);
        });
    }
}

const axiosJWTDecoratorInstance = new AxiosJWTDecorator();
export default axiosJWTDecoratorInstance;