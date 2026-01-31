
import { ApiResponse, TravelRequest, CostCenterMaster, SupportData, Integrant } from '../types';
import { API_BASE_URL } from '../constants';

class GasService {
  
  /**
   * Universal Bridge using HTTP FETCH.
   */
  private async runGas(action: string, payload: any = null): Promise<ApiResponse<any>> {
    if (!API_BASE_URL || API_BASE_URL.includes('REPLACE')) {
      console.error("API URL not configured in constants.ts");
      return { success: false, error: "API URL no configurada. Revise constants.ts" };
    }

    try {
      // Using mode: 'cors' and redirect: 'follow' is standard for GAS Web Apps.
      // We use text/plain content-type to avoid preflight OPTIONS requests which GAS doesn't handle natively for simple triggers.
      const response = await fetch(API_BASE_URL, {
        method: 'POST',
        mode: 'cors', 
        redirect: 'follow',
        headers: {
          'Content-Type': 'text/plain;charset=utf-8', 
        },
        body: JSON.stringify({
          action,
          payload
        })
      });

      if (!response.ok) {
        throw new Error(`HTTP Error: ${response.status} ${response.statusText}`);
      }

      const textResult = await response.text();
      let result;
      try {
          result = JSON.parse(textResult);
      } catch (e) {
          console.error("Invalid JSON response:", textResult);
          throw new Error("El servidor devolvió una respuesta inválida.");
      }

      return result;

    } catch (error: any) {
      console.error(`[API Error] ${action}:`, error);
      
      let msg = error.message || 'Error de conexión';
      if (msg.includes('Failed to fetch')) {
          msg = 'No se pudo conectar con el servidor. Verifique su conexión o la URL del script.';
      }
      return { success: false, error: msg };
    }
  }

  async getCurrentUser(): Promise<string> {
    const response = await this.runGas('getCurrentUser');
    return response.data || '';
  }

  async getMyRequests(userEmail: string): Promise<TravelRequest[]> {
    const response = await this.runGas('getMyRequests', { userEmail });
    return response.data || [];
  }

  async getAllRequests(userEmail: string): Promise<TravelRequest[]> {
    const response = await this.runGas('getAllRequests', { userEmail });
    return response.data || [];
  }

  async createRequest(request: Partial<TravelRequest>): Promise<string> {
    const response = await this.runGas('createRequest', request);
    if (!response.success) throw new Error(response.error);
    return response.data;
  }

  async updateRequestStatus(id: string, status: string, payload?: any): Promise<void> {
    const response = await this.runGas('updateRequest', { id, status, payload });
    if (!response.success) throw new Error(response.error);
  }
  
  async getCostCenterData(): Promise<CostCenterMaster[]> {
    const response = await this.runGas('getCostCenterData');
    return response.data || [];
  }

  async getIntegrantesData(): Promise<Integrant[]> {
    const response = await this.runGas('getIntegrantesData');
    return response.data || [];
  }

  async uploadSupportFile(requestId: string, fileData: string, fileName: string, mimeType: string): Promise<SupportData> {
    const response = await this.runGas('uploadSupportFile', { requestId, fileData, fileName, mimeType });
    if (!response.success) throw new Error(response.error);
    return response.data;
  }

  async closeRequest(requestId: string): Promise<void> {
    const response = await this.runGas('closeRequest', { requestId });
    if (!response.success) throw new Error(response.error);
  }
}

export const gasService = new GasService();
