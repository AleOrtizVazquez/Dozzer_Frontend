import { HttpErrorResponse, HttpInterceptorFn } from '@angular/common/http';
import { catchError, throwError } from 'rxjs';

// La sesión dura lo que indica el servidor (token JWT). Si ya venció, se avisa a la app
// para que pida entrar de nuevo, en lugar de quedarse mostrando datos viejos sin poder guardar.
export function tokenVencido(token: string | null): boolean {
 if (!token) return false;
 try {
  const parte = token.split('.')[1] || '';
  const json = JSON.parse(atob(parte.replace(/-/g, '+').replace(/_/g, '/')));
  return typeof json.exp === 'number' && json.exp * 1000 < Date.now();
 } catch {
  return false;
 }
}

function avisarSesionVencida(): void {
 try { window.dispatchEvent(new Event('dozzer-sesion-vencida')); } catch { /* */ }
}

export const authInterceptor: HttpInterceptorFn = (req, next) => {
 const token = localStorage.getItem('dozzer_token');
 const esLogin = req.url.includes('/auth/');
 if (token && !esLogin && tokenVencido(token) && navigator.onLine !== false) avisarSesionVencida();
 return next(token ? req.clone({ setHeaders: { Authorization: `Bearer ${token}` } }) : req).pipe(
  catchError((e: HttpErrorResponse) => {
   if (e.status === 401 && token && !esLogin) avisarSesionVencida();
   return throwError(() => e);
  })
 );
};
