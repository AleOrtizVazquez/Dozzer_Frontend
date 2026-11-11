
import { Component, ElementRef, ViewChild } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { HttpClient, HttpHeaders } from '@angular/common/http';
import * as XLSX from 'xlsx';
import * as QRCode from 'qrcode';
import { API_URL, QR_FICHA_URL } from './config';
import { jsPDF } from 'jspdf';
import jsQR from 'jsqr';
import { DomSanitizer, SafeResourceUrl } from '@angular/platform-browser';
import { firstValueFrom } from 'rxjs';

// Ventana de confirmación con diseño propio
interface DialogoConfirmacion {
  tipo: 'peligro' | 'confirmar';
  icono: string;
  titulo: string;
  mensaje: string;
  textoConfirmar: string;
  textoCancelar?: string;
  nota?: string;
  bien?: any;
  cambios?: { etiqueta: string; actual: string; propuesto: string }[];
  cargando?: boolean;
  error?: string;
  accion: (terminar: (error?: string) => void) => void;
}

// Datos para llenar los formatos oficiales FA-04-02-03 (Alta) y FA-04-02-04 (Baja)
interface FormatoDoc {
  tipo: 'ALTA' | 'BAJA' | 'ACTA_INICIO' | 'ACTA_CIERRE' | 'RESGUARDO';
  folio: string;
  sede: string;
  fecha: string;
  filas: { inventario: string; descripcion: string; modeloSerie: string }[];
  observaciones: string;
  solicito: string;
  autorizo: string;
  detalle?: { etiqueta: string; valor: string }[];
  correoSugerido?: string;
  contactos?: { nombre: string; correo: string; rol: string }[];
  anulado?: boolean;
  borrador?: boolean;
  historico?: boolean;
  // Resguardo individual
  persona?: { nombre?: string; numeroEmpleado?: string; area?: string; puesto?: string };
  totalBienes?: number;
}

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, FormsModule],
  template: `
  <div class="app-shell">

    <!-- MENÚ LATERAL -->
    <div class="sidebar-backdrop" *ngIf="logged && !mustChangePassword && !menuCollapsed" (click)="menuCollapsed=true"></div>
    <aside class="sidebar" [class.collapsed]="menuCollapsed" *ngIf="logged && !mustChangePassword">
      <div class="brand">
        <div class="brand-icon">D</div>
        <div class="brand-text">
          <strong>DOZZER</strong>
          <small>Construcción y Asesoría S.A. de C.V.</small>
        </div>
        <button class="sidebar-close" title="Cerrar menú" (click)="menuCollapsed=true">×</button>
      </div>

      <div class="menu-label">MENÚ PRINCIPAL</div>

      <button class="nav-item" [class.active]="section==='inicio'" (click)="menuCollapsed=true; section='inicio'; editing=false">
        <span class="nav-icon">⌂</span>
        <span>Inicio</span>
      </button>

      <button class="nav-item" [class.active]="section==='bienes'" (click)="menuCollapsed=true; section='bienes'; vistaInv='menu'; editing=false">
        <span class="nav-icon">▦</span>
        <span>Bienes</span>
      </button>

      <button class="nav-item" *ngIf="verInventarioAnual" [class.active]="section==='anual'" (click)="menuCollapsed=true; section='anual'; anualVista='lista'; loadLevantamientos()">
        <span class="nav-icon">✓</span>
        <span>Levantamiento</span>
      </button>

      <button class="nav-item" [class.active]="section==='solicitudes'" (click)="menuCollapsed=true; section='solicitudes'; loadSolicitudes()">
        <span class="nav-icon">▣</span>
        <span>{{ isAdmin ? 'Autorizaciones' : 'Solicitudes' }}</span>
        <span class="nav-badge" *ngIf="pendientesBadge > 0">{{ pendientesBadge }}</span>
      </button>

      <button class="nav-item" [class.active]="section==='documentos'" (click)="menuCollapsed=true; section='documentos'; docsVista='menu'; loadDocumentos()">
        <span class="nav-icon">▤</span>
        <span>Documentos</span>
      </button>

      <button class="nav-item" *ngIf="isAdmin" [class.active]="section==='catalogos'" (click)="menuCollapsed=true; section='catalogos'; loadCatalog()">
        <span class="nav-icon">⚙</span>
        <span>Catálogos</span>
      </button>

      <div class="sidebar-bottom" *ngIf="logged">
        <div class="user-avatar">{{ initials }}</div>
        <div class="user-info">
          <strong [title]="email">{{ nombreUsuario || email }}</strong>
          <small>{{ isAdmin ? 'Administrador' : 'Empleado' }}</small>
        </div>
        <button class="logout" title="Cambiar contraseña" (click)="abrirCambioPassword()">⚿</button>
        <button class="logout" title="Cerrar sesión" (click)="logout()">↪</button>
      </div>
    </aside>

    <!-- CONTENIDO -->
    <main class="main-content">

      <!-- LOGIN -->
      <section class="lg-screen" *ngIf="!logged">
        <form class="lg-card" (ngSubmit)="doLogin()" autocomplete="on">

          <img class="lg-logo-img" src="logo-dozzer.png" alt="DOZZER Construcción y Asesoría S.A. de C.V.">

          <div class="lg-titles">
            <h1>INVENTARIOS</h1>
            <p>CONTROL &amp; ACTIVOS</p>
          </div>

          <ng-container *ngIf="!modoRecuperar">
          <div class="lg-field" [class.invalid]="error">
            <input id="lg-correo" name="correo" [(ngModel)]="login.correo" type="email" placeholder="CORREO ELECTRÓNICO"
                   aria-label="Correo electrónico" autocomplete="username" (input)="error=''">
          </div>

          <div class="lg-field" [class.invalid]="error">
            <input id="lg-pass" name="password" [(ngModel)]="login.password" [type]="verPasswordLogin ? 'text' : 'password'"
                   placeholder="CONTRASEÑA" aria-label="Contraseña" autocomplete="current-password" (input)="error=''">
            <button type="button" class="lg-eye" (click)="verPasswordLogin = !verPasswordLogin" [title]="verPasswordLogin ? 'Ocultar contraseña' : 'Mostrar contraseña'">
              <svg *ngIf="!verPasswordLogin" viewBox="0 0 24 24"><path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12z"/><circle cx="12" cy="12" r="3"/></svg>
              <svg *ngIf="verPasswordLogin" viewBox="0 0 24 24"><path d="M3 3l18 18M10.6 5.1A10 10 0 0 1 12 5c6.5 0 10 7 10 7a17 17 0 0 1-3.2 4.1M6.6 6.6A17 17 0 0 0 2 12s3.5 7 10 7a9.6 9.6 0 0 0 5.4-1.6M9.9 9.9a3 3 0 0 0 4.2 4.2"/></svg>
            </button>
          </div>

          <div class="lg-error" *ngIf="error">
            <svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"/><path d="M12 8v5M12 16h.01"/></svg>
            <span>{{ error }}</span>
          </div>

          <button type="submit" class="lg-submit" [disabled]="loading">
            <span class="mini-ring light" *ngIf="loading"></span>
            {{ loading ? 'ENTRANDO…' : 'ENTRAR' }}
          </button>

          <button type="button" class="lg-forgot" (click)="abrirRecuperacion()">¿Olvidaste tu contraseña?</button>

          </ng-container>

          <!-- RECUPERAR CONTRASEÑA -->
          <ng-container *ngIf="modoRecuperar">
            <div class="lg-recover">
              <h2>Restablecer contraseña</h2>
              <p>Escribe tu correo y tu número de empleado, y crea una contraseña nueva. Al guardarla entrarás directamente.</p>
            </div>

            <div class="lg-field" [class.invalid]="recuperacionError">
              <input name="correoRecuperar" [(ngModel)]="correoRecuperar" type="email" placeholder="CORREO ELECTRÓNICO"
                     aria-label="Correo electrónico" autocomplete="username" (input)="recuperacionError=''">
            </div>
            <div class="lg-field" [class.invalid]="recuperacionError">
              <input name="numeroRecuperar" [(ngModel)]="numeroRecuperar" type="text" placeholder="NÚMERO DE EMPLEADO (EJ. DZ-0001)"
                     aria-label="Número de empleado" autocomplete="off" (input)="recuperacionError=''">
            </div>
            <div class="lg-field" [class.invalid]="recuperacionError">
              <input name="nuevaRecuperar" [(ngModel)]="nuevaRecuperar" [type]="verPasswordLogin ? 'text' : 'password'"
                     placeholder="NUEVA CONTRASEÑA" aria-label="Nueva contraseña" autocomplete="new-password" (input)="recuperacionError=''">
              <button type="button" class="lg-eye" (click)="verPasswordLogin = !verPasswordLogin" [title]="verPasswordLogin ? 'Ocultar contraseña' : 'Mostrar contraseña'">
                <svg *ngIf="!verPasswordLogin" viewBox="0 0 24 24"><path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12z"/><circle cx="12" cy="12" r="3"/></svg>
                <svg *ngIf="verPasswordLogin" viewBox="0 0 24 24"><path d="M3 3l18 18M10.6 5.1A10 10 0 0 1 12 5c6.5 0 10 7 10 7a17 17 0 0 1-3.2 4.1M6.6 6.6A17 17 0 0 0 2 12s3.5 7 10 7a9.6 9.6 0 0 0 5.4-1.6M9.9 9.9a3 3 0 0 0 4.2 4.2"/></svg>
              </button>
            </div>
            <div class="lg-field" [class.invalid]="recuperacionError">
              <input name="confirmarRecuperar" [(ngModel)]="confirmarRecuperar" [type]="verPasswordLogin ? 'text' : 'password'"
                     placeholder="CONFIRMAR CONTRASEÑA" aria-label="Confirmar contraseña" autocomplete="new-password" (input)="recuperacionError=''"
                     (keydown.enter)="$event.preventDefault(); solicitarRecuperacion()">
            </div>

            <ul class="lg-rules">
              <li [class.ok]="(nuevaRecuperar || '').length >= 8">Mínimo 8 caracteres</li>
              <li [class.ok]="tieneLetrasYNumeros(nuevaRecuperar)">Letras y números</li>
              <li [class.ok]="nuevaRecuperar && nuevaRecuperar === confirmarRecuperar">Las dos coinciden</li>
            </ul>

            <div class="lg-error" *ngIf="recuperacionError">
              <svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"/><path d="M12 8v5M12 16h.01"/></svg>
              <span>{{ recuperacionError }}</span>
            </div>

            <button type="button" class="lg-submit" [disabled]="recuperacionEnviando" (click)="solicitarRecuperacion()">
              <span class="mini-ring light" *ngIf="recuperacionEnviando"></span>
              {{ recuperacionEnviando ? 'GUARDANDO…' : 'GUARDAR Y ENTRAR' }}
            </button>

            <button type="button" class="lg-forgot" (click)="cerrarRecuperacion()">← Volver a iniciar sesión</button>
          </ng-container>

          <footer class="lg-footer">
            <p>DOZZER INVENTORY CONTROL SYSTEM V1.0 • © {{ anioActual }}</p>          </footer>
        </form>
      </section>

      <!-- CAMBIO OBLIGATORIO DE CONTRASEÑA (PRIMER INICIO DE SESIÓN) -->
      <section class="login-screen" *ngIf="logged && mustChangePassword">
        <div class="login-card password-card">
          <div class="brand login-brand">
            <div class="brand-icon">D</div>
            <strong>DOZZER</strong>
          </div>

          <h1>Crea tu contraseña</h1>
          <p>Hola{{ nombreUsuario ? ', ' + nombreUsuario : '' }}. Entraste con una contraseña temporal; por tu seguridad, crea una contraseña personal para continuar.</p>

          <ng-container *ngTemplateOutlet="formPassword"></ng-container>

          <button class="btn-primary full" [disabled]="pwdGuardando || !pwdValida" (click)="guardarPassword()">
            {{ pwdGuardando ? 'Guardando...' : 'Guardar y continuar' }}
          </button>
          <button class="link-button" [disabled]="pwdGuardando" (click)="logout()">Cerrar sesión</button>
        </div>
      </section>

      <!-- CAMBIO VOLUNTARIO DE CONTRASEÑA -->
      <div class="modal-backdrop" *ngIf="cambioPasswordModal">
        <div class="modal-card password-modal">
          <div class="modal-header">
            <div>
              <h2>Cambiar contraseña</h2>
              <p>Actualiza la contraseña con la que entras a DOZZER.</p>
            </div>
            <button class="close-button" [disabled]="pwdGuardando" (click)="cambioPasswordModal=false">×</button>
          </div>
          <div class="form-section">
            <ng-container *ngTemplateOutlet="formPassword"></ng-container>
          </div>
          <div class="modal-footer">
            <button class="btn-outline" [disabled]="pwdGuardando" (click)="cambioPasswordModal=false">Cancelar</button>
            <button class="btn-primary" [disabled]="pwdGuardando || !pwdValida" (click)="guardarPassword()">
              {{ pwdGuardando ? 'Guardando...' : 'Guardar contraseña' }}
            </button>
          </div>
        </div>
      </div>

      <!-- FORMATO PDF DE ALTA / BAJA -->
      <div class="modal-backdrop doc-backdrop" *ngIf="documento || docGenerando">
        <div class="modal-card doc-card">
          <div class="doc-loading" *ngIf="docGenerando && !documento">
            <div class="loader-ring"></div>
            <strong>Preparando el formato…</strong>
          </div>

          <ng-container *ngIf="documento as d">
            <div class="modal-header">
              <div class="doc-title">
                <span class="doc-badge" [class.baja]="d.datos.tipo==='BAJA'" [class.acta]="esActa(d.datos)" [class.resg]="d.datos.tipo==='RESGUARDO'">{{ etiquetaTipoDoc(d.datos) }}</span>
                <div>
                  <h2>{{ tituloDoc(d.datos) }} <span class="doc-anulado" *ngIf="d.datos.anulado">ANULADO</span><span class="doc-estado borrador" *ngIf="d.datos.borrador">BORRADOR</span><span class="doc-estado" *ngIf="d.datos.historico">HISTÓRICO</span></h2>
                  <p>Folio {{ d.datos.folio }}<ng-container *ngIf="!esActa(d.datos)"> · {{ d.datos.tipo==='ALTA' ? 'FA-04-02-03' : d.datos.tipo==='BAJA' ? 'FA-04-02-04' : 'FA-04-02-02' }}</ng-container> · {{ d.datos.tipo==='RESGUARDO' ? d.datos.persona?.nombre : (d.datos.sede || 'Sin sede') }}</p>
                </div>
              </div>
              <button class="close-button" (click)="cerrarFormato()">×</button>
            </div>

            <!-- Vista principal: vista previa + acciones -->
            <div class="doc-body" *ngIf="vistaDoc==='inicio'">
              <div class="doc-preview">
                <iframe [src]="d.vista" title="Vista previa del formato"></iframe>
              </div>

              <div class="doc-side">
                <h3>¿Qué quieres hacer?</h3>
                <button class="doc-action" (click)="descargarFormato()">
                  <span>⬇</span><div><strong>Descargar PDF</strong><small>Guardarlo en este equipo</small></div>
                </button>
                <button class="doc-action" (click)="imprimirFormato()">
                  <span>⎙</span><div><strong>Abrir / imprimir</strong><small>Para firmarlo en papel</small></div>
                </button>
                <button class="doc-action" (click)="abrirVistaDoc('whatsapp')">
                  <span class="ico-wa"><svg viewBox="0 0 24 24"><path d="M12 2a10 10 0 0 0-8.6 15.1L2 22l5-1.3A10 10 0 1 0 12 2zm0 18.2a8.2 8.2 0 0 1-4.2-1.2l-.3-.2-3 .8.8-2.9-.2-.3A8.2 8.2 0 1 1 12 20.2zm4.5-6.1c-.2-.1-1.5-.7-1.7-.8s-.4-.1-.6.1-.6.8-.8 1-.3.2-.5.1a6.7 6.7 0 0 1-3.3-2.9c-.3-.4.3-.4.7-1.3.1-.2 0-.3 0-.4l-.8-1.8c-.2-.5-.4-.4-.6-.4h-.5a1 1 0 0 0-.7.3 3 3 0 0 0-.9 2.2 5.2 5.2 0 0 0 1.1 2.7 11.8 11.8 0 0 0 4.5 4c1.7.7 2.3.8 3.2.6a2.7 2.7 0 0 0 1.8-1.3 2.2 2.2 0 0 0 .2-1.3c-.1-.1-.3-.2-.5-.3z"/></svg></span><div><strong>Compartir por WhatsApp</strong><small>Abre la app con el mensaje listo</small></div><em>›</em>
                </button>
                <div class="doc-file-mini">📄 {{ d.nombre }}</div>
              </div>
            </div>

            <!-- Vista: WhatsApp -->
            <div class="doc-flow" *ngIf="vistaDoc==='whatsapp'">
              <button class="doc-back" (click)="vistaDoc='inicio'">← Volver al formato</button>
              <div class="flow-head">
                <span class="flow-ico ico-wa"><svg viewBox="0 0 24 24"><path d="M12 2a10 10 0 0 0-8.6 15.1L2 22l5-1.3A10 10 0 1 0 12 2zm0 18.2a8.2 8.2 0 0 1-4.2-1.2l-.3-.2-3 .8.8-2.9-.2-.3A8.2 8.2 0 1 1 12 20.2zm4.5-6.1c-.2-.1-1.5-.7-1.7-.8s-.4-.1-.6.1-.6.8-.8 1-.3.2-.5.1a6.7 6.7 0 0 1-3.3-2.9c-.3-.4.3-.4.7-1.3.1-.2 0-.3 0-.4l-.8-1.8c-.2-.5-.4-.4-.6-.4h-.5a1 1 0 0 0-.7.3 3 3 0 0 0-.9 2.2 5.2 5.2 0 0 0 1.1 2.7 11.8 11.8 0 0 0 4.5 4c1.7.7 2.3.8 3.2.6a2.7 2.7 0 0 0 1.8-1.3 2.2 2.2 0 0 0 .2-1.3c-.1-.1-.3-.2-.5-.3z"/></svg></span>
                <div>
                  <h2>Compartir por WhatsApp</h2>
                  <p>Archivo: <b>{{ d.nombre }}</b></p>
                </div>
              </div>

              <div class="flow-step">
                <div class="step-title"><span>1</span> Número de WhatsApp <em>(opcional)</em></div>
                <input class="wa-input" [(ngModel)]="waNumero" type="tel" inputmode="numeric" maxlength="16" placeholder="10 dígitos, ej. 4611234567" (input)="waAviso=''">
                <small class="flow-help">Si lo dejas vacío, WhatsApp te pedirá elegir el contacto.</small>
              </div>

              <div class="flow-step">
                <div class="step-title"><span>2</span> Así funciona</div>
                <ol class="wa-steps">
                  <li>Se abre WhatsApp con el mensaje ya escrito.</li>
                  <li>El PDF se descarga en tu equipo.</li>
                  <li>En el chat toca <b>📎 → Documento</b> y adjúntalo.</li>
                </ol>
              </div>

              <div class="wa-hint" *ngIf="waAviso">
                {{ waAviso }}
                <a *ngIf="waUrl" class="wa-manual" [href]="waUrl" target="_blank" rel="noopener">¿No se abrió WhatsApp? Toca aquí</a>
              </div>
              <div class="flow-actions">
                <button class="btn-outline" *ngIf="puedeCompartir" (click)="compartirFormato()">Enviar el archivo directo (celular)</button>
                <button class="wa-btn" (click)="abrirWhatsApp()"><svg viewBox="0 0 24 24"><path d="M12 2a10 10 0 0 0-8.6 15.1L2 22l5-1.3A10 10 0 1 0 12 2zm0 18.2a8.2 8.2 0 0 1-4.2-1.2l-.3-.2-3 .8.8-2.9-.2-.3A8.2 8.2 0 1 1 12 20.2zm4.5-6.1c-.2-.1-1.5-.7-1.7-.8s-.4-.1-.6.1-.6.8-.8 1-.3.2-.5.1a6.7 6.7 0 0 1-3.3-2.9c-.3-.4.3-.4.7-1.3.1-.2 0-.3 0-.4l-.8-1.8c-.2-.5-.4-.4-.6-.4h-.5a1 1 0 0 0-.7.3 3 3 0 0 0-.9 2.2 5.2 5.2 0 0 0 1.1 2.7 11.8 11.8 0 0 0 4.5 4c1.7.7 2.3.8 3.2.6a2.7 2.7 0 0 0 1.8-1.3 2.2 2.2 0 0 0 .2-1.3c-.1-.1-.3-.2-.5-.3z"/></svg> Abrir WhatsApp</button>
              </div>
            </div>
          </ng-container>
        </div>
      </div>

      <!-- RESTABLECER MANUAL (ADMIN) -->
      <div class="modal-backdrop" *ngIf="restablecerManualAbierto">
        <div class="modal-card" style="max-width:460px">
          <div class="modal-header">
            <div>
              <h2>Restablecer contraseña</h2>
              <p>Asigna una contraseña temporal a cualquier usuario, aunque no la haya solicitado.</p>
            </div>
            <button class="close-button" (click)="restablecerManualAbierto=false">×</button>
          </div>
          <div class="form-section">
            <label>Correo del usuario</label>
            <input [(ngModel)]="correoRestablecer" type="email" placeholder="nombre.apellido@dozzer.mx"
                   (keyup.enter)="correoRestablecer && restablecerPassword(correoRestablecer)">
          </div>
          <div class="modal-footer">
            <button class="btn-outline" (click)="restablecerManualAbierto=false">Cancelar</button>
            <button class="btn-primary" [disabled]="!correoRestablecer.trim()" (click)="restablecerPassword(correoRestablecer)">Continuar</button>
          </div>
        </div>
      </div>

      <!-- CONTRASEÑA TEMPORAL GENERADA -->
      <div class="modal-backdrop" *ngIf="passwordTemporal" style="z-index:210">
        <div class="modal-card temp-card" style="max-width:440px">
          <div class="temp-ico">✓</div>
          <h2>Contraseña temporal lista</h2>
          <p>Para <b>{{ passwordTemporal.nombre || passwordTemporal.correo }}</b><br>{{ passwordTemporal.correo }}</p>
          <div class="temp-pass">
            <code>{{ passwordTemporal.passwordTemporal }}</code>
            <button class="btn-outline btn-small" (click)="copiarTemporal()">{{ temporalCopiada ? '✓ Copiada' : 'Copiar' }}</button>
          </div>
          <ol class="temp-steps">
            <li>Entrégasela a la persona en privado.</li>
            <li>Al iniciar sesión, el sistema le pedirá crear su propia contraseña.</li>
            <li>Por seguridad, <b>no se volverá a mostrar</b> al cerrar esta ventana.</li>
          </ol>
          <button class="btn-primary full" (click)="passwordTemporal=null">Listo</button>
        </div>
      </div>

      <!-- AVISO DE RESULTADO (creado, modificado, eliminado...) -->
      <div class="aviso-backdrop" *ngIf="aviso" (click)="cerrarAviso()">
        <div class="aviso-card" [ngClass]="'av-' + aviso.tipo" (click)="$event.stopPropagation()" role="status" aria-live="polite">
          <div class="aviso-ico">
            <svg *ngIf="aviso.tipo==='exito'" viewBox="0 0 52 52"><circle class="av-circ" cx="26" cy="26" r="23"/><path class="av-trazo" d="M15 27l7 7 15-16"/></svg>
            <svg *ngIf="aviso.tipo==='pendiente'" viewBox="0 0 52 52"><circle class="av-circ" cx="26" cy="26" r="23"/><path class="av-trazo" d="M26 14v13l8 5"/></svg>
            <svg *ngIf="aviso.tipo==='baja'" viewBox="0 0 52 52"><circle class="av-circ" cx="26" cy="26" r="23"/><path class="av-trazo" d="M17 18h18M23 18v-3h6v3M19 18l1.5 19h11L33 18M24 23v9M28 23v9"/></svg>
            <svg *ngIf="aviso.tipo==='rechazo' || aviso.tipo==='error'" viewBox="0 0 52 52"><circle class="av-circ" cx="26" cy="26" r="23"/><path class="av-trazo" d="M18 18l16 16M34 18L18 34"/></svg>
            <svg *ngIf="aviso.tipo==='info'" viewBox="0 0 52 52"><circle class="av-circ" cx="26" cy="26" r="23"/><path class="av-trazo" d="M26 24v12M26 16v.5"/></svg>
          </div>
          <h2>{{ aviso.titulo }}</h2>
          <p>{{ aviso.mensaje }}</p>
          <ng-container *ngIf="aviso.accion; else soloEntendido">
            <button class="aviso-btn" (click)="ejecutarAccionAviso()">{{ aviso.accion.texto }}</button>
            <button class="aviso-btn-sec" (click)="cerrarAviso()">Ahora no</button>
          </ng-container>
          <ng-template #soloEntendido>
            <button class="aviso-btn" (click)="cerrarAviso()">Entendido</button>
          </ng-template>
          <span class="aviso-timer" *ngIf="aviso.ms"><span [style.animation-duration.ms]="aviso.ms"></span></span>
        </div>
      </div>

      <!-- DIÁLOGO DE CONFIRMACIÓN -->
      <div class="confirm-backdrop" *ngIf="dialogo as d" (click)="cerrarDialogo()">
        <div class="confirm-card" [ngClass]="'confirm-' + d.tipo" (click)="$event.stopPropagation()" role="alertdialog" aria-modal="true">
          <div class="confirm-icon">{{ d.icono }}</div>
          <h2>{{ d.titulo }}</h2>
          <p>{{ d.mensaje }}</p>

          <div class="confirm-asset" *ngIf="d.bien">
            <span class="asset-key">{{ d.bien.clave }}</span>
            <strong>{{ d.bien.articulo }}</strong>
            <small *ngIf="d.bien.ubicacion">{{ d.bien.ubicacion }}</small>
          </div>

          <div class="confirm-changes" *ngIf="d.cambios?.length">
            <div class="confirm-changes-title">{{ d.cambios?.length }} {{ d.cambios?.length === 1 ? 'campo modificado' : 'campos modificados' }}</div>
            <div class="change-row" *ngFor="let c of d.cambios">
              <span class="change-label">{{ c.etiqueta }}</span>
              <span class="change-old">{{ c.actual }}</span>
              <span class="change-arrow">→</span>
              <span class="change-new">{{ c.propuesto }}</span>
            </div>
          </div>

          <div class="confirm-note" *ngIf="d.nota">{{ d.nota }}</div>
          <div class="inline-error" *ngIf="d.error">{{ d.error }}</div>

          <div class="confirm-actions">
            <button class="btn-outline" [disabled]="d.cargando" (click)="cerrarDialogo()">{{ d.textoCancelar || 'Cancelar' }}</button>
            <button [ngClass]="d.tipo === 'peligro' ? 'btn-danger-solid' : 'btn-primary'" [disabled]="d.cargando" (click)="confirmarDialogo()">
              <span class="mini-ring light" *ngIf="d.cargando"></span>
              {{ d.cargando ? 'Procesando...' : d.textoConfirmar }}
            </button>
          </div>
        </div>
      </div>

      <!-- FORMULARIO DE CONTRASEÑA (compartido) -->
      <ng-template #formPassword>
        <div class="password-form">
          <label>{{ mustChangePassword ? 'Contraseña temporal' : 'Contraseña actual' }}</label>
          <input [(ngModel)]="pwd.actual" type="password" autocomplete="current-password" (input)="pwdError=''">

          <label>Nueva contraseña</label>
          <input [(ngModel)]="pwd.nueva" type="password" autocomplete="new-password" (input)="pwdError=''">

          <label>Confirmar nueva contraseña</label>
          <input [(ngModel)]="pwd.confirmar" type="password" autocomplete="new-password" (input)="pwdError=''" (keyup.enter)="guardarPassword()">

          <ul class="password-rules">
            <li [class.ok]="pwdReglas.longitud">Al menos 8 caracteres</li>
            <li [class.ok]="pwdReglas.letrasNumeros">Incluye letras y números</li>
            <li [class.ok]="pwdReglas.distinta">Distinta a la contraseña actual</li>
            <li [class.ok]="pwdReglas.coincide">Las contraseñas coinciden</li>
          </ul>

          <div class="error-message" *ngIf="pwdError">{{ pwdError }}</div>
        </div>
      </ng-template>

      <ng-container *ngIf="logged && !mustChangePassword">

        <!-- TOPBAR -->
        <header class="topbar">
          <button class="menu-toggle" title="Menú" (click)="menuCollapsed=!menuCollapsed">☰<span class="toggle-badge" *ngIf="pendientesBadge > 0">{{ pendientesBadge }}</span></button>
          <div class="topbar-title">
            <small>DOZZER / {{ sectionLabel }}</small>
            <strong>{{ sectionTitle }}</strong>
          </div>
          <div class="topbar-right">
            <span class="online-dot"></span>
            <span class="online-label">Sistema activo</span>
            <div class="user-menu-wrap">
              <button class="top-avatar top-avatar-btn" [class.open]="menuUsuario" title="Mi cuenta" (click)="menuUsuario=!menuUsuario">{{ initials }}</button>
              <div class="picker-backdrop" *ngIf="menuUsuario" (click)="menuUsuario=false"></div>
              <div class="user-menu" *ngIf="menuUsuario">
                <div class="user-menu-head">
                  <span class="top-avatar">{{ initials }}</span>
                  <div>
                    <strong>{{ nombreUsuario || email }}</strong>
                    <small>{{ isAdmin ? 'Administrador' : 'Empleado' }} · {{ email }}</small>
                  </div>
                </div>
                <button *ngIf="isAdmin" (click)="menuUsuario=false; abrirRestablecerManual()">
                  <span class="more-ico">⚿</span>
                  <div><strong>Restablecer contraseña</strong><small>Asignar una temporal a un usuario</small></div>
                </button>
                <button *ngIf="isAdmin" (click)="menuUsuario=false; generarTemporalesPendientes()">
                  <span class="more-ico">⇩</span>
                  <div><strong>Contraseñas temporales</strong><small>Una distinta por empleado (Excel)</small></div>
                </button>
                <button *ngIf="instalarPrompt" (click)="menuUsuario=false; instalarApp()">
                  <span class="more-ico">📲</span>
                  <div><strong>Instalar DOZZER</strong><small>Como app en este dispositivo</small></div>
                </button>
                <button [disabled]="pushProcesando" (click)="menuUsuario=false; pushEstado==='activo' ? desactivarPush() : activarPush()">
                  <span class="more-ico">{{ pushEstado==='activo' ? '🔔' : '🔕' }}</span>
                  <div><strong>{{ pushEstado==='activo' ? 'Avisos activados' : 'Activar avisos' }}</strong><small>{{ textoPushEstado }}</small></div>
                </button>
                <button (click)="menuUsuario=false; abrirCambioPassword()">
                  <span class="more-ico">✎</span>
                  <div><strong>Cambiar mi contraseña</strong><small>Actualiza tu acceso</small></div>
                </button>
                <button class="user-menu-exit" (click)="menuUsuario=false; logout()">
                  <span class="more-ico">↪</span>
                  <div><strong>Cerrar sesión</strong></div>
                </button>
              </div>
            </div>
          </div>
        </header>

        <div class="content-area">



          <!-- INICIO: SALUDO -->
          <section class="welcome" *ngIf="section==='inicio'">
            <div class="welcome-inner">
              <img class="welcome-logo" src="logo-dozzer.png" alt="DOZZER Construcción y Asesoría S.A. de C.V.">
              <span class="welcome-chip">{{ saludoHora }} · {{ fechaHoy }}</span>
              <h1>Hola, {{ primerNombre }}</h1>
              <p>Bienvenido al sistema oficial de inventarios.<br>Selecciona una opción del menú lateral para comenzar.</p>
              <div class="push-invita" *ngIf="(pushEstado==='inactivo' || pushEstado==='instalar-ios') && !avisoPushOculto">
                <span class="push-ico">🔔</span>
                <div class="push-txt">
                  <strong>Recibe avisos en este celular</strong>
                  <small>{{ isAdmin || canAuthorize ? 'Que te suene cuando haya autorizaciones pendientes.' : 'Te avisamos cuando aprueben o rechacen tus solicitudes.' }}</small>
                </div>
                <button class="btn-primary btn-small" [disabled]="pushProcesando" (click)="activarPush()">Activar</button>
                <button class="push-x" title="Ahora no" (click)="ocultarAvisoPush()">×</button>
              </div>
              <button class="push-instalar" *ngIf="instalarPrompt" (click)="instalarApp()">📲 Instalar DOZZER como app</button>
              <div class="cm-aviso" *ngIf="(isAdmin || canAuthorize) && alertasCM.total as total">
                <span class="push-ico">🛠</span>
                <div class="push-txt">
                  <strong>Calibración y mantenimiento</strong>
                  <small>
                    <ng-container *ngIf="alertasCM.vencidas"><b class="cm-n vencida">{{ alertasCM.vencidas }} vencidos</b> · </ng-container>
                    <ng-container *ngIf="alertasCM.proximas"><b class="cm-n proxima">{{ alertasCM.proximas }} en 30 días</b> · </ng-container>
                    {{ alertasCM.sinFecha }} sin fecha de {{ total }}
                  </small>
                </div>
                <button class="btn-outline btn-small" (click)="verAlertasCM(alertasCM.vencidas || alertasCM.proximas ? 'atencion' : 'sinfecha')">Ver</button>
              </div>
            </div>
            <footer class="welcome-foot">DOZZER INVENTORY CONTROL SYSTEM V1.0 • © {{ anioActual }}</footer>
          </section>

          <!-- INVENTARIO -->
          <section *ngIf="section==='bienes'">

            <div class="inv-head">
              <div>
                <button class="inv-back" *ngIf="vistaInv!=='menu'" (click)="volverInventario()">← {{ textoVolverInv }}</button>
                <h1>{{ tituloVistaInv }}</h1>
                <p class="inv-resumen" *ngIf="vistaInv!=='menu'">{{ subtituloVistaInv }}</p>
                <ng-container *ngIf="vistaInv==='menu'">
                <p class="inv-resumen" *ngIf="bienes.length; else resumenVacio">
                  <span><b>{{ bienes.length | number }}</b> bienes</span>
                  <span><b>{{ totalPiezas | number }}</b> piezas</span>
                  <span><b>{{ porcentaje(bienesBuenEstado) }}%</b> en buen estado</span>
                  <span><b>{{ bienesDisponibles | number }}</b> sin asignar</span>
                  <span class="inv-scope">{{ isAdmin ? 'Todas las sedes' : 'Tu sede' }}</span>
                </p>
                <ng-template #resumenVacio><p class="inv-resumen">{{ bienesCargando ? 'Cargando…' : 'Bienes y equipo de la empresa' }}</p></ng-template>
                </ng-container>
              </div>
              <div class="heading-actions">
                <div class="more-wrap">
                  <button class="btn-outline" (click)="menuMasInventario=!menuMasInventario">
                    Más opciones <span class="caret">▾</span>
                  </button>
                  <div class="picker-backdrop" *ngIf="menuMasInventario" (click)="menuMasInventario=false"></div>
                  <div class="more-menu" *ngIf="menuMasInventario">
                    <button *ngIf="isAdmin" (click)="menuMasInventario=false; abrirEtiquetas()">
                      <span class="more-ico">🏷</span>
                      <div><strong>Etiquetas QR en blanco</strong><small>Calcular, generar e imprimir hojas</small></div>
                    </button>
                    <button *ngIf="isAdmin" (click)="menuMasInventario=false; openExcel()">
                      <span class="more-ico">▤</span>
                      <div><strong>Importar desde Excel</strong><small>Carga masiva de bienes</small></div>
                    </button>
                  </div>
                </div>
                <button class="btn-primary" (click)="newBien()">
                  <span>＋</span> Nuevo registro
                </button>
              </div>
            </div>

            <!-- CARGA INICIAL: ESQUELETOS ANIMADOS -->
            <ng-container *ngIf="bienesCargando && !bienes.length">
              <div class="inventory-panel skeleton-panel">
                <div class="loading-banner">
                  <div class="loader-ring"></div>
                  <div>
                    <strong>Preparando tu inventario</strong>
                    <p>Estamos reuniendo la información de tus bienes. Esto puede tardar unos segundos.</p>
                  </div>
                </div>
                <div class="sk-chips"><span class="sk sk-chip" *ngFor="let i of [1,2,3,4,5,6]"></span></div>
                <div class="sk-row" *ngFor="let i of [1,2,3,4,5,6,7,8]">
                  <span class="sk sk-line w10"></span><span class="sk sk-line w30"></span>
                  <span class="sk sk-line w20"></span><span class="sk sk-line w15"></span><span class="sk sk-pill"></span>
                </div>
              </div>
            </ng-container>

            <!-- ERROR DE CARGA -->
            <div class="inventory-panel placeholder-panel" *ngIf="bienesError && !bienes.length && !bienesCargando">
              <div class="placeholder-icon">⚠</div>
              <h2>No fue posible cargar el inventario</h2>
              <p>{{ bienesError }}</p>
              <button class="btn-primary" (click)="loadBienes()">Reintentar</button>
            </div>

            <ng-container *ngIf="bienes.length || (!bienesCargando && !bienesError)">
              <div class="fade-in">

                <!-- MENÚ DE INVENTARIO: TARJETAS GRANDES -->
                <div class="inv-tiles" *ngIf="vistaInv==='menu'">
                  <button class="inv-tile t-sedes" (click)="abrirVistaInv('sedes')">
                    <svg class="tile-bg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M3 21h18M5 21V7l7-4 7 4v14M9 9h1M14 9h1M9 13h1M14 13h1M10 21v-4h4v4"/></svg>
                    <span class="tile-ico"><svg class="" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M3 21h18M5 21V7l7-4 7 4v14M9 9h1M14 9h1M9 13h1M14 13h1M10 21v-4h4v4"/></svg></span>
                    <span class="tile-text">
                      <strong>Sedes</strong>
                      <small>Bienes agrupados por sede</small>
                    </span>
                    <span class="tile-num">{{ resumenSedes.length | number }}<em>{{ resumenSedes.length === 1 ? 'sede' : 'sedes' }}</em></span>
                    <span class="tile-go">Ver sedes →</span>
                  </button>
                  <button class="inv-tile t-bienes" (click)="abrirVistaInv('bienes')">
                    <svg class="tile-bg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M21 8l-9-5-9 5 9 5 9-5z"/><path d="M3 8v8l9 5 9-5V8"/><path d="M12 13v8"/></svg>
                    <span class="tile-ico"><svg class="" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M21 8l-9-5-9 5 9 5 9-5z"/><path d="M3 8v8l9 5 9-5V8"/><path d="M12 13v8"/></svg></span>
                    <span class="tile-text">
                      <strong>Bienes</strong>
                      <small>Cada bien en una tarjeta</small>
                    </span>
                    <span class="tile-num">{{ bienes.length | number }}<em>bienes</em></span>
                    <span class="tile-go">Ver bienes →</span>
                  </button>
                  <button class="inv-tile t-personas" (click)="abrirVistaInv('personas')">
                    <svg class="tile-bg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="9" cy="8" r="3.5"/><path d="M2.5 20c.8-3.6 3.4-5.5 6.5-5.5s5.7 1.9 6.5 5.5"/><circle cx="17" cy="9" r="2.6"/><path d="M16.5 14.6c2.6.2 4.4 1.9 5 5.4"/></svg>
                    <span class="tile-ico"><svg class="" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="9" cy="8" r="3.5"/><path d="M2.5 20c.8-3.6 3.4-5.5 6.5-5.5s5.7 1.9 6.5 5.5"/><circle cx="17" cy="9" r="2.6"/><path d="M16.5 14.6c2.6.2 4.4 1.9 5 5.4"/></svg></span>
                    <span class="tile-text">
                      <strong>Personas</strong>
                      <small>Quién tiene cada bien a su cargo</small>
                    </span>
                    <span class="tile-num">{{ resumenPersonas.length | number }}<em>personas</em></span>
                    <span class="tile-go">Ver personas →</span>
                  </button>
                  <button class="inv-tile t-completo" (click)="abrirVistaInv('completo')">
                    <svg class="tile-bg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 9h18M3 14h18M9 9v11M15 9v11"/></svg>
                    <span class="tile-ico"><svg class="" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 9h18M3 14h18M9 9v11M15 9v11"/></svg></span>
                    <span class="tile-text">
                      <strong>Completo</strong>
                      <small>Tabla con búsqueda, filtros y columnas</small>
                    </span>
                    <span class="tile-num">{{ totalPiezas | number }}<em>piezas</em></span>
                    <span class="tile-go">Abrir tabla →</span>
                  </button>
                </div>

                <!-- VISTA: SEDES -->
                <div class="sede-grid" *ngIf="vistaInv==='sedes'">
                  <button class="sede-card" *ngFor="let s of resumenSedes" (click)="verBienesDeSede(s)">
                    <div class="sede-top">
                      <span class="sede-ico"><svg class="" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M3 21h18M5 21V7l7-4 7 4v14M9 9h1M14 9h1M9 13h1M14 13h1M10 21v-4h4v4"/></svg></span>
                      <div>
                        <strong>{{ s.nombre }}</strong>
                        <small>{{ s.ubicaciones | number }} {{ s.ubicaciones === 1 ? 'ubicación' : 'ubicaciones' }} · {{ s.personas | number }} personas</small>
                      </div>
                    </div>
                    <div class="sede-num">{{ s.total | number }}<span>bienes</span></div>
                    <div class="sede-bar"><span [style.width.%]="s.pctBueno"></span></div>
                    <div class="sede-meta">
                      <span>{{ s.piezas | number }} piezas</span>
                      <span>{{ s.pctBueno }}% en buen estado</span>
                    </div>
                    <div class="sede-tags">
                      <span *ngFor="let c of s.top"><i [style.background]="colorDeClasificacion(c.id)"></i>{{ c.nombre | titlecase }} {{ c.total | number }}</span>
                    </div>
                    <span class="sede-go">Ver bienes →</span>
                  </button>
                  <div class="inv-empty" *ngIf="!resumenSedes.length">No hay bienes registrados en ninguna sede.</div>
                </div>

                <!-- VISTA: PERSONAS -->
                <ng-container *ngIf="vistaInv==='personas'">
                  <div class="people-bar">
                    <div class="search-box">
                      <span>⌕</span>
                      <input [(ngModel)]="busquedaPersonas" (ngModelChange)="personasVisibles=48" placeholder="Buscar persona por nombre...">
                      <button class="clear-search" *ngIf="busquedaPersonas" (click)="busquedaPersonas=''">×</button>
                    </div>
                    <span class="people-count">{{ personasFiltradas.length | number }} {{ personasFiltradas.length === 1 ? 'persona' : 'personas' }}</span>
                  </div>
                  <div class="people-grid">
                    <button class="person-card" *ngFor="let p of personasFiltradas.slice(0, personasVisibles)" (click)="verBienesDePersona(p)">
                      <span class="person-avatar">{{ inicialesDe(p.nombre) }}</span>
                      <span class="person-info">
                        <strong>{{ p.nombre }}</strong>
                        <small>{{ p.sedes || 'Sin sede' }}</small>
                        <span class="person-tags">
                          <span *ngIf="p.responsable">{{ p.responsable }} como responsable</span>
                          <span *ngIf="p.usuario">{{ p.usuario }} como usuario</span>
                        </span>
                      </span>
                      <span class="person-num">{{ p.total | number }}<em>{{ p.total === 1 ? 'bien' : 'bienes' }}</em></span>
                    </button>
                  </div>
                  <div class="inv-empty" *ngIf="!personasFiltradas.length">{{ busquedaPersonas ? 'Ninguna persona coincide con la búsqueda.' : 'Ningún bien tiene responsable o usuario asignado.' }}</div>
                  <div class="more-cards" *ngIf="personasFiltradas.length > personasVisibles">
                    <button class="btn-outline" (click)="personasVisibles = personasVisibles + 48">Mostrar más personas ({{ personasFiltradas.length - personasVisibles | number }} restantes)</button>
                  </div>
                </ng-container>

                <!-- BIENES EN TARJETAS / TABLA COMPLETA -->
                <div class="inventory-panel" [class.cards-mode]="vistaInv==='bienes'" *ngIf="vistaInv==='bienes' || vistaInv==='completo'">

                  <!-- Clasificaciones como pestañas -->
                  <div class="class-tabs" *ngIf="resumenClasificaciones.length">
                    <button [class.active]="!filtroClasificacion" (click)="setFiltroClasificacion(null)">
                      Todas <span>{{ bienes.length | number }}</span>
                    </button>
                    <button *ngFor="let c of resumenClasificaciones; let i = index" [class.active]="filtroClasificacion===c.id"
                            (click)="setFiltroClasificacion(filtroClasificacion===c.id ? null : c.id)" [title]="c.pct + '% del inventario'">
                      <i [style.background]="colorClasificacion(i)"></i>{{ c.nombre | titlecase }} <span>{{ c.total | number }}</span>
                    </button>
                  </div>

                  <div class="inv-toolbar">
                    <div class="search-box">
                      <span>⌕</span>
                      <input [(ngModel)]="filter" (ngModelChange)="buscarBienes()" placeholder="Buscar por clave, bien, serie, responsable o ubicación…">
                      <button class="clear-search" *ngIf="filter" (click)="filter=''; aplicarFiltrosBienes()">×</button>
                    </div>
                    <div class="table-actions">
                      <span class="refreshing" *ngIf="bienesCargando"><span class="mini-ring"></span> Actualizando…</span>
                      <div class="column-picker">
                        <button class="btn-outline" [class.filters-active]="filtrosActivos.length" (click)="columnPickerOpen=!columnPickerOpen">
                          <span>⚲</span> Filtros
                          <span class="tab-count" *ngIf="filtrosActivos.length">{{ filtrosActivos.length }}</span>
                        </button>
                        <div class="picker-backdrop" *ngIf="columnPickerOpen" (click)="columnPickerOpen=false"></div>
                        <div class="picker-menu filters-menu" *ngIf="columnPickerOpen">
                          <div class="picker-head">
                            <strong>Filtrar bienes</strong>
                            <small>Los resultados se actualizan al instante.</small>
                          </div>
                          <div class="filter-fields">
                            <label *ngIf="isAdmin">Sede
                              <select [(ngModel)]="filtroSede" (ngModelChange)="aplicarFiltrosBienes()">
                                <option [ngValue]="null">Todas las sedes</option>
                                <option *ngFor="let s of activos(catalogs.sedes, filtroSede, 'id_sede')" [ngValue]="s.id_sede">{{ s.nombre_sede }}</option>
                              </select>
                            </label>
                            <label>Clasificación
                              <select [(ngModel)]="filtroClasificacion" (ngModelChange)="aplicarFiltrosBienes()">
                                <option [ngValue]="null">Todas las clasificaciones</option>
                                <option *ngFor="let c of catalogs.clasificaciones" [ngValue]="c.id_clasificacion">{{ c.nombre }}</option>
                              </select>
                            </label>
                            <label>Estatus
                              <select [(ngModel)]="statusFilter" (ngModelChange)="aplicarFiltrosBienes()">
                                <option value="">Todos los estatus</option>
                                <option value="ACTIVO">Activo</option>
                                <option value="BAJA">Baja</option>
                                <option value="EN_REPARACION">En reparación</option>
                              </select>
                            </label>
                            <label>Estado físico
                              <select [(ngModel)]="filtroEstadoFisico" (ngModelChange)="aplicarFiltrosBienes()">
                                <option value="">Todos</option>
                                <option value="BUENO">Bueno</option>
                                <option value="REGULAR">Regular</option>
                                <option value="MALO">Malo</option>
                              </select>
                            </label>
                            <label>Asignación
                              <select [(ngModel)]="filtroAsignacion" (ngModelChange)="aplicarFiltrosBienes()">
                                <option value="">Todos</option>
                                <option value="ASIGNADO">Asignados</option>
                                <option value="SIN_ASIGNAR">Sin asignar</option>
                              </select>
                            </label>
                            <label>Calibración y mantenimiento
                              <select [(ngModel)]="filtroCM" (ngModelChange)="aplicarFiltrosBienes()">
                                <option value="">Todos</option>
                                <option value="atencion">⚠ Vencidos o próximos (30 días)</option>
                                <option value="calibracion">🟢 Requieren calibración</option>
                                <option value="mantenimiento">🔵 Requieren mantenimiento</option>
                                <option value="sinfecha">Sin fecha asignada</option>
                              </select>
                            </label>
                          </div>

                          <ng-container *ngIf="vistaInv==='completo'">
                          <div class="picker-head picker-subhead">
                            <strong>Columnas visibles</strong>
                            <small>Elige qué información muestra la tabla.</small>
                          </div>
                          <div class="picker-presets">
                            <button (click)="aplicarVistaColumnas('basica')">Básica</button>
                            <button (click)="aplicarVistaColumnas('ubicacion')">Ubicación</button>
                            <button (click)="aplicarVistaColumnas('completa')">Completa</button>
                          </div>
                          <div class="picker-list">
                            <label *ngFor="let c of columnasDisponibles" [class.locked]="c.fija">
                              <input type="checkbox" [checked]="columnaVisible(c.key)" [disabled]="c.fija" (change)="toggleColumna(c.key)">
                              {{ c.label }} <small *ngIf="c.fija">fija</small>
                            </label>
                          </div>

                          </ng-container>

                          <div class="picker-footer">
                            <button class="link-button inline" [disabled]="!filtrosActivos.length" (click)="limpiarFiltrosBienes()">Limpiar filtros</button>
                            <button class="btn-primary btn-small" (click)="columnPickerOpen=false">Ver {{ bienesVista.length | number }} bienes</button>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>

                  <div class="search-row inventory-filters" *ngIf="filtrosActivos.length">
                    <div class="active-filters">
                      <span class="filter-chip" *ngFor="let f of filtrosActivos">
                        {{ f.etiqueta }}: <b>{{ f.valor }}</b>
                        <button title="Quitar filtro" (click)="quitarFiltro(f.clave)">×</button>
                      </span>
                    </div>
                    <button class="link-button inline" *ngIf="hayFiltrosBienes" (click)="limpiarFiltrosBienes()">Limpiar filtros</button>
                  </div>

                  <!-- Bienes como tarjetas -->
                  <ng-container *ngIf="vistaInv==='bienes'">
                    <div class="asset-cards">
                      <button class="asset-card" *ngFor="let b of bienesVista.slice(0, tarjetasVisibles); trackBy: trackBien"
                              [style.border-left-color]="colorDeClasificacion(b.id_clasificacion)" (click)="abrirBien(b)">
                        <span class="asset-chips">
                          <span class="chip-key">{{ b.clave || 'SIN CLAVE' }}</span>
                          <span class="chip-class" *ngIf="b.clasificacion">{{ b.clasificacion }}</span>
                          <span class="status" [ngClass]="statusClass(b)">{{ displayStatus(b) }}</span>
                        </span>
                        <strong class="asset-title">{{ b.articulo || 'Bien sin nombre' }}</strong>
                        <small class="asset-sub" *ngIf="marcaModelo(b)">{{ marcaModelo(b) }}</small>
                        <span class="asset-line"><svg class="" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 21s-7-6.2-7-11.5A7 7 0 0 1 19 9.5C19 14.8 12 21 12 21z"/><circle cx="12" cy="9.5" r="2.5"/></svg>{{ b.ubicacion || 'Sin ubicación' }}<ng-container *ngIf="isAdmin && b.sede"> · {{ b.sede }}</ng-container></span>
                        <span class="asset-line" [class.muted]="!b.responsable"><svg class="" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="8" r="4"/><path d="M4 21c1.5-4 4.5-6 8-6s6.5 2 8 6"/></svg>{{ b.responsable || 'Sin responsable' }}</span>
                        <span class="asset-foot">
                          <span>{{ b.cantidad_piezas || 1 }} {{ (b.cantidad_piezas || 1) == 1 ? 'pieza' : 'piezas' }}</span>
                          <span class="asset-more">Ver detalle →</span>
                        </span>
                      </button>
                    </div>
                    <div class="inv-empty" *ngIf="!bienesVista.length">
                      {{ hayFiltrosBienes ? 'Ningún bien coincide con los filtros.' : 'Aún no hay bienes registrados.' }}
                      <button class="link-button inline" *ngIf="hayFiltrosBienes" (click)="limpiarFiltrosBienes()">Limpiar filtros</button>
                    </div>
                    <div class="more-cards" *ngIf="bienesVista.length">
                      <span>Mostrando {{ (bienesVista.length < tarjetasVisibles ? bienesVista.length : tarjetasVisibles) | number }} de {{ bienesVista.length | number }} bienes</span>
                      <button class="btn-outline" *ngIf="bienesVista.length > tarjetasVisibles" (click)="tarjetasVisibles = tarjetasVisibles + 36">Mostrar más</button>
                    </div>
                  </ng-container>

                  <ng-container *ngIf="vistaInv==='completo'">
                  <div class="table-container">
                    <table class="inventory-table">
                      <thead>
                        <tr>
                          <th *ngFor="let c of columnasVisibles" class="sortable" [class.num]="c.tipo==='num'" (click)="ordenarBienes(c.key)">
                            {{ c.label | uppercase }}
                            <span class="sort-icon" [class.active]="ordenBienes.key===c.key">{{ ordenBienes.key===c.key ? (ordenBienes.dir==='asc' ? '▲' : '▼') : '↕' }}</span>
                          </th>
                          <th class="chevron-col"></th>
                        </tr>
                      </thead>
                      <tbody>
                        <tr *ngFor="let b of paginaBienes; trackBy: trackBien" class="clickable-row" [class.row-selected]="bienSeleccionado?.id_bien===b.id_bien" (click)="abrirBien(b)">
                          <td *ngFor="let c of columnasVisibles" [class.num]="c.tipo==='num'" [ngSwitch]="c.tipo">
                            <strong *ngSwitchCase="'clave'" class="asset-key">{{ b.clave }}</strong>
                            <div *ngSwitchCase="'bien'" class="asset-name">
                              <strong>{{ b.articulo }}</strong>
                              <small>{{ marcaModelo(b) }}</small>
                            </div>
                            <span *ngSwitchCase="'num'" class="quantity">{{ b[c.key] || 1 }}</span>
                            <span *ngSwitchCase="'estatus'" class="status" [ngClass]="statusClass(b)">{{ displayStatus(b) }}</span>
                            <span *ngSwitchCase="'fisico'" class="status" [ngClass]="claseEstadoFisico(b.estado_fisico)">{{ b.estado_fisico || '—' }}</span>
                            <span *ngSwitchCase="'fecha'">{{ b[c.key] ? (b[c.key] | date:'dd/MM/yyyy') : '—' }}</span>
                            <span *ngSwitchCase="'persona'" [class.muted]="!b[c.key]">{{ b[c.key] || 'Sin asignar' }}</span>
                            <span *ngSwitchDefault [class.muted]="!b[c.key]">{{ textoCampo(b[c.key]) }}</span>
                          </td>
                          <td class="chevron-col">›</td>
                        </tr>

                        <tr *ngIf="!bienesVista.length">
                          <td [attr.colspan]="columnasVisibles.length + 1" class="empty-state">
                            <div class="empty-icon">▦</div>
                            <strong>{{ hayFiltrosBienes ? 'Ningún bien coincide con los filtros' : 'Aún no hay bienes registrados' }}</strong>
                            <p>{{ hayFiltrosBienes ? 'Prueba con otra búsqueda o limpia los filtros.' : 'Registra el primer bien para comenzar.' }}</p>
                            <button class="btn-outline" *ngIf="hayFiltrosBienes" (click)="limpiarFiltrosBienes()">Limpiar filtros</button>
                            <button class="btn-primary" *ngIf="!hayFiltrosBienes" (click)="newBien()">Nuevo registro</button>
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </div>

                  <div class="table-footer">
                    <span>
                      Mostrando {{ rangoBienes.desde | number }}–{{ rangoBienes.hasta | number }} de {{ bienesVista.length | number }} bienes
                      <ng-container *ngIf="bienesVista.length !== bienes.length"> (filtrados de {{ bienes.length | number }})</ng-container>
                    </span>
                    <div class="pager">
                      <select class="page-size" [(ngModel)]="tamanoPaginaBienes" (ngModelChange)="paginaActualBienes=1">
                        <option [ngValue]="25">25 por página</option>
                        <option [ngValue]="50">50 por página</option>
                        <option [ngValue]="100">100 por página</option>
                      </select>
                      <button class="btn-outline btn-small" [disabled]="paginaActualBienes<=1" (click)="paginaActualBienes=paginaActualBienes-1">‹</button>
                      <span>Página {{ paginaActualBienes }} de {{ totalPaginasBienes }}</span>
                      <button class="btn-outline btn-small" [disabled]="paginaActualBienes>=totalPaginasBienes" (click)="paginaActualBienes=paginaActualBienes+1">›</button>
                    </div>
                  </div>
                  </ng-container>
                </div>
              </div>
            </ng-container>

            <!-- ETIQUETAS QR EN BLANCO -->
            <div class="modal-backdrop" *ngIf="etq">
              <div class="modal-card wizard">
                <div class="modal-header">
                  <div>
                    <h2>🏷 Etiquetas QR en blanco</h2>
                    <p>Hoja carta adhesiva de 30 etiquetas (tipo Avery 5160). Se pegan en cualquier bien y se ligan al escanearlas.</p>
                  </div>
                  <button class="close-button" [disabled]="etq.generando" (click)="etq=null">×</button>
                </div>
                <div class="form-section">
                  <div class="cap-loading" *ngIf="etq.cargando"><div class="loader-ring"></div><strong>Calculando…</strong></div>
                  <ng-container *ngIf="!etq.cargando && etq.resumen as r">
                    <div class="etq-stats">
                      <div><b>{{ r.total | number }}</b><small>Impresas</small></div>
                      <div><b>{{ r.asignadas | number }}</b><small>Pegadas</small></div>
                      <div><b>{{ r.libres | number }}</b><small>Sin pegar</small></div>
                    </div>

                    <button class="lev-more" (click)="etq.verSedes = !etq.verSedes">{{ etq.verSedes ? '▴' : '▾' }} Bienes sin etiqueta por sede ({{ etqFaltan | number }})</button>
                    <div class="etq-sedes" *ngIf="etq.verSedes">
                      <div *ngFor="let x of r.porSede"><span>{{ x.sede }}</span><b>{{ x.sin_etiqueta | number }}</b><small>≈ {{ hojasPara(x.sin_etiqueta) }} hojas</small></div>
                    </div>

                    <label>¿Cuántas hojas vas a imprimir?</label>
                    <div class="etq-hojas">
                      <button class="btn-outline" (click)="etq.hojas = mathMax(1, etq.hojas - 1)">−</button>
                      <input type="number" min="1" max="30" [(ngModel)]="etq.hojas">
                      <button class="btn-outline" (click)="etq.hojas = mathMin(30, etq.hojas + 1)">＋</button>
                    </div>
                    <p class="cap-hint">
                      = {{ etq.hojas * 30 | number }} etiquetas, de <b>{{ codigoEtiqueta(r.ultimo + 1) }}</b> a <b>{{ codigoEtiqueta(r.ultimo + etq.hojas * 30) }}</b>.
                      Para terminar todo faltan ≈ <b>{{ hojasSugeridas | number }} hojas</b> (incluye 5% extra); imprímelas por partes de hasta 30.
                    </p>
                    <div class="wiz-summary">En la impresora elige <b>tamaño carta</b>, <b>escala 100%</b> y <b>sin márgenes</b>. Haz primero una prueba en hoja normal y empálmala con la adhesiva.</div>
                    <button class="lev-more" *ngIf="r.libres" [disabled]="etq.generando" (click)="reimprimirLibres()">⎙ Volver a imprimir las {{ mathMin(r.libres, 900) | number }} que siguen sin pegar</button>
                  </ng-container>
                </div>
                <div class="inline-error" *ngIf="etq.error" style="margin:0 24px 12px">{{ etq.error }}</div>
                <div class="modal-footer">
                  <button class="btn-outline" [disabled]="etq.generando" (click)="etq=null">Cerrar</button>
                  <button class="btn-primary" [disabled]="etq.cargando || etq.generando || !(etq.hojas >= 1)" (click)="generarEtiquetas()">
                    <span class="mini-ring light" *ngIf="etq.generando"></span>
                    {{ etq.generando ? 'Generando…' : '⎙ Generar e imprimir ' + (etq.hojas || 0) + (etq.hojas === 1 ? ' hoja' : ' hojas') }}
                  </button>
                </div>
              </div>
            </div>

            <!-- ASIGNAR ETIQUETA A UN BIEN (desde su ficha) -->
            <div class="modal-backdrop" *ngIf="etqAsignar">
              <div class="modal-card wizard">
                <div class="modal-header">
                  <div>
                    <h2>🏷 Asignar etiqueta</h2>
                    <p><b>{{ etqAsignar.bien.clave }}</b> · {{ etqAsignar.bien.articulo }}</p>
                  </div>
                  <button class="close-button" [disabled]="etqAsignar.guardando" (click)="etqAsignar=null">×</button>
                </div>
                <div class="form-section">
                  <label>Código de la etiqueta que pegaste</label>
                  <div class="cap-clave">
                    <input [(ngModel)]="etqAsignar.codigo" placeholder="Ej. DZ-000123" (keyup.enter)="confirmarAsignarEtiqueta()">
                    <label class="btn-outline etq-foto">📷<input type="file" accept="image/*" capture="environment" (change)="leerEtiquetaDeFoto($event)" hidden></label>
                  </div>
                  <p class="cap-hint">Escribe el código impreso o toma foto del QR.</p>
                  <label class="cm-check" *ngIf="etqAsignar.bien.etiqueta"><input type="checkbox" [(ngModel)]="etqAsignar.reemplazar"> Reemplazar la etiqueta {{ etqAsignar.bien.etiqueta }} (quedará anulada)</label>
                </div>
                <div class="inline-error" *ngIf="etqAsignar.error" style="margin:0 24px 12px">{{ etqAsignar.error }}</div>
                <div class="modal-footer">
                  <button class="btn-outline" [disabled]="etqAsignar.guardando" (click)="etqAsignar=null">Cancelar</button>
                  <button class="btn-primary" [disabled]="etqAsignar.guardando || !etqAsignar.codigo" (click)="confirmarAsignarEtiqueta()">
                    <span class="mini-ring light" *ngIf="etqAsignar.guardando"></span> Asignar
                  </button>
                </div>
              </div>
            </div>

            <!-- PANEL LATERAL: DETALLE Y ACCIONES DEL BIEN -->
            <div class="drawer-backdrop" *ngIf="bienSeleccionado" (click)="bienSeleccionado=null"></div>
            <aside class="drawer" *ngIf="bienSeleccionado as b">
              <div class="drawer-header">
                <div>
                  <span class="asset-key">{{ b.clave }}</span>
                  <h2>{{ b.articulo }}</h2>
                  <div class="drawer-badges">
                    <span class="status" [ngClass]="statusClass(b)">{{ displayStatus(b) }}</span>
                    <span class="status" *ngIf="b.estado_fisico" [ngClass]="claseEstadoFisico(b.estado_fisico)">Estado {{ b.estado_fisico | lowercase }}</span>
                    <span class="status status-neutral">{{ b.estatus_asignacion === 'ASIGNADO' ? 'Asignado' : 'Sin asignar' }}</span>
                  </div>
                </div>
                <button class="close-button" (click)="bienSeleccionado=null">×</button>
              </div>

              <div class="drawer-body">
                <div class="drawer-section drawer-foto" *ngIf="b.foto_mini">
                  <button type="button" class="foto-mini" (click)="ampliarFoto(b)" title="Ver foto en grande">
                    <img [src]="b.foto_mini" [alt]="'Foto de ' + b.articulo" loading="lazy">
                    <span>🔍 Ver en grande</span>
                  </button>
                </div>
                <div class="drawer-section">
                  <h3>Información general</h3>
                  <dl>
                    <div><dt>Clasificación</dt><dd>{{ textoCampo(b.clasificacion) }}</dd></div>
                    <div><dt>Subcategoría</dt><dd>{{ textoCampo(b.subcategoria) }}</dd></div>
                    <div><dt>Marca</dt><dd>{{ textoCampo(b.marca) }}</dd></div>
                    <div><dt>Modelo</dt><dd>{{ textoCampo(b.modelo) }}</dd></div>
                    <div><dt>Número de serie</dt><dd>{{ textoCampo(b.numero_serie) }}</dd></div>
                    <div><dt>Color</dt><dd>{{ textoCampo(b.color) }}</dd></div>
                    <div><dt>Piezas</dt><dd>{{ b.cantidad_piezas || 1 }} {{ b.unidad_medida || '' }}</dd></div>
                    <div><dt>Fecha de alta</dt><dd>{{ b.fecha_alta ? (b.fecha_alta | date:'dd/MM/yyyy') : '—' }}</dd></div>
                  </dl>
                </div>

                <div class="drawer-section">
                  <h3>Ubicación y responsables</h3>
                  <dl>
                    <div><dt>Sede</dt><dd>{{ textoCampo(b.sede) }}</dd></div>
                    <div><dt>Ubicación</dt><dd>{{ textoCampo(b.ubicacion) }}</dd></div>
                    <div><dt>Responsable de área</dt><dd>{{ b.responsable || 'Sin asignar' }}</dd></div>
                    <div><dt>Usuario directo</dt><dd>{{ b.usuario_directo || 'Sin asignar' }}</dd></div>
                  </dl>
                </div>

                <div class="drawer-section" *ngIf="b.requiere_calibracion || b.requiere_mantenimiento">
                  <h3>Calibración y mantenimiento</h3>
                  <div class="cm-chips">
                    <span *ngIf="b.requiere_calibracion" class="cm-chip" [ngClass]="'cm-' + estadoFechaCM(b.fecha_calibracion)">🟢 Calibración · {{ textoFechaCM(b.fecha_calibracion) }}</span>
                    <span *ngIf="b.requiere_mantenimiento" class="cm-chip" [ngClass]="'cm-' + estadoFechaCM(b.fecha_mantenimiento)">🔵 Mantenimiento · {{ textoFechaCM(b.fecha_mantenimiento) }}</span>
                  </div>
                </div>

                <div class="drawer-section" *ngIf="b.caracteristicas">
                  <h3>Características</h3>
                  <div class="note-box">{{ b.caracteristicas }}</div>
                </div>

                <div class="drawer-section">
                  <h3>Etiqueta QR</h3>
                  <div class="qr-block">
                    <div class="qr-frame">
                      <img *ngIf="qrPreview" [src]="qrPreview" alt="Código QR del bien">
                      <span *ngIf="!qrPreview && b.etiqueta" class="mini-ring"></span>
                      <span *ngIf="!b.etiqueta" class="etq-vacia">🏷</span>
                    </div>
                    <div class="qr-info">
                      <ng-container *ngIf="b.etiqueta">
                        <small>Etiqueta pegada</small>
                        <strong class="qr-id">{{ b.etiqueta }}</strong>
                        <p>El QR solo lleva este código. Al escanearlo con DOZZER se abre este bien.</p>
                        <div class="etq-btns">
                          <button class="btn-outline btn-small" (click)="imprimirCodigos([b.etiqueta])">⎙ Reimprimir</button>
                          <button class="btn-outline btn-small" *ngIf="isAdmin" (click)="abrirAsignarEtiqueta(b)">Cambiar</button>
                        </div>
                      </ng-container>
                      <ng-container *ngIf="!b.etiqueta">
                        <small>Sin etiqueta</small>
                        <strong class="qr-id">—</strong>
                        <p>Pega una etiqueta en blanco y asígnala aquí, o escanéala durante el levantamiento.</p>
                        <button class="btn-outline btn-small" *ngIf="isAdmin" (click)="abrirAsignarEtiqueta(b)">🏷 Asignar etiqueta</button>
                      </ng-container>
                    </div>
                  </div>
                </div>

                <div class="info-box" *ngIf="!isAdmin">
                  <strong>Los cambios requieren autorización</strong>
                  <p>Al editar o dar de baja este bien se enviará una solicitud al administrador.</p>
                </div>
              </div>

              <div class="drawer-footer">
                <button class="btn-outline" *ngIf="isAdmin" (click)="formatoAltaDeBien(b)">📄 Formato de alta</button>
                <button class="btn-danger" (click)="deleteBien(b.id_bien)">⌫ {{ isAdmin ? 'Dar de baja' : 'Solicitar baja' }}</button>
                <button class="btn-primary" (click)="editarDesdePanel(b)">✎ {{ isAdmin ? 'Editar bien' : 'Solicitar cambio' }}</button>
              </div>
            </aside>

            <!-- FORMULARIO DE BIEN -->
            <div class="modal-backdrop" *ngIf="editing">
              <div class="modal-card large-modal">
                <div class="modal-header">
                  <div>
                    <h2>{{ bien.idBien ? 'Editar registro' : 'Nuevo registro' }}</h2>
                    <p>Completa la información del bien.</p>
                  </div>
                  <button class="close-button" (click)="editing=false">×</button>
                </div>

                <div class="form-section">
                  <h3>Información general</h3>
                  <div class="form-grid">
                    <div>
                      <label>Clave manual *</label>
                      <input [(ngModel)]="bien.clave" placeholder="Ej. ED0001">
                    </div>
                    <div>
                      <label>Nombre del bien *</label>
                      <input [(ngModel)]="bien.articulo" placeholder="Ej. Laptop">
                    </div>
                    <div>
                      <label>Clasificación</label>
                      <select [(ngModel)]="bien.idClasificacion" (ngModelChange)="onClasificacionBienChange()">
                        <option [ngValue]="null">Seleccionar</option>
                        <option *ngFor="let x of catalogs.clasificaciones" [ngValue]="x.id_clasificacion">{{ x.nombre }}</option>
                      </select>
                    </div>
                    <div>
                      <label>Subcategoría</label>
                      <select [(ngModel)]="bien.idSubcategoria" (ngModelChange)="onSubcategoriaBienChange()">
                        <option [ngValue]="null">{{ subcategoriasBien.length ? 'Seleccionar' : 'Sin subcategorías' }}</option>
                        <option *ngFor="let x of subcategoriasBien" [ngValue]="x.id_subcategoria">{{ x.nombre_subcategoria }}</option>
                      </select>
                    </div>
                    <div>
                      <label>Marca</label>
                      <select [(ngModel)]="bien.idMarca">
                        <option [ngValue]="null">Seleccionar</option>
                        <option *ngFor="let x of catalogs.marcas" [ngValue]="x.id_marca">{{ x.nombre_marca }}</option>
                      </select>
                    </div>
                    <div>
                      <label>Modelo</label>
                      <input [(ngModel)]="bien.modelo">
                    </div>
                    <div>
                      <label>Número de serie</label>
                      <input [(ngModel)]="bien.numeroSerie">
                    </div>
                    <div>
                      <label>Color</label>
                      <select [(ngModel)]="bien.idColor">
                        <option [ngValue]="null">Seleccionar</option>
                        <option *ngFor="let x of catalogs.colores" [ngValue]="x.id_color">{{ x.nombre_color }}</option>
                      </select>
                    </div>
                    <div>
                      <label>Unidad de medida</label>
                      <input [(ngModel)]="bien.unidadMedida" placeholder="Pieza">
                    </div>
                    <div>
                      <label>Cantidad de piezas</label>
                      <input type="number" min="1" [(ngModel)]="bien.cantidadPiezas">
                    </div>
                    <div>
                      <label>Responsable de área</label>
                      <select [(ngModel)]="bien.idResponsableArea">
                        <option [ngValue]="null">Seleccionar</option>
                        <option *ngFor="let x of activos(catalogs.personas, bien.idResponsableArea, 'id_persona')" [ngValue]="x.id_persona">{{ x.nombre_completo }}</option>
                      </select>
                    </div>
                    <div>
                      <label>Usuario directo</label>
                      <select [(ngModel)]="bien.idUsuarioDirecto">
                        <option [ngValue]="null">Seleccionar</option>
                        <option *ngFor="let x of activos(catalogs.personas, bien.idUsuarioDirecto, 'id_persona')" [ngValue]="x.id_persona">{{ x.nombre_completo }}</option>
                      </select>
                    </div>
                    <div>
                      <label>Ubicación</label>
                      <select [(ngModel)]="bien.idUbicacion">
                        <option [ngValue]="null">Seleccionar</option>
                        <option *ngFor="let x of activos(catalogs.ubicaciones, bien.idUbicacion, 'id_ubicacion')" [ngValue]="x.id_ubicacion">{{ x.nombre_area }}{{ x.nombre_sede && isAdmin ? ' · ' + x.nombre_sede : '' }}</option>
                      </select>
                    </div>
                    <div>
                      <label>Estado físico</label>
                      <select [(ngModel)]="bien.estadoFisico">
                        <option value="">Seleccionar</option>
                        <option>BUENO</option>
                        <option>REGULAR</option>
                        <option>MALO</option>
                      </select>
                    </div>
                    <div>
                      <label>Estatus de inventario</label>
                      <select [(ngModel)]="bien.estatusInventario">
                        <option>ACTIVO</option>
                        <option>BAJA</option>
                        <option>EN_REPARACION</option>
                      </select>
                    </div>
                    <div>
                      <label>Estatus de asignación</label>
                      <select [(ngModel)]="bien.estatusAsignacion">
                        <option>SIN_ASIGNAR</option>
                        <option>ASIGNADO</option>
                      </select>
                    </div>
                  </div>
                  <div class="cm-form">
                    <div class="cm-row">
                      <label class="cm-check"><input type="checkbox" [(ngModel)]="bien.requiereCalibracion"> 🟢 Requiere calibración</label>
                      <input *ngIf="bien.requiereCalibracion" type="date" [(ngModel)]="bien.fechaCalibracion" title="Próxima calibración">
                    </div>
                    <div class="cm-row">
                      <label class="cm-check"><input type="checkbox" [(ngModel)]="bien.requiereMantenimiento"> 🔵 Requiere mantenimiento</label>
                      <input *ngIf="bien.requiereMantenimiento" type="date" [(ngModel)]="bien.fechaMantenimiento" title="Próximo mantenimiento">
                    </div>
                  </div>
                  <label>Características y descripción</label>
                  <textarea [(ngModel)]="bien.caracteristicas" rows="3" placeholder="Describe las características del bien..."></textarea>
                </div>

                <div class="modal-footer">
                  <button class="btn-outline" (click)="editing=false">Cancelar</button>
                  <button class="btn-primary" [disabled]="saving" (click)="saveBien()">
                    {{ saving ? 'Guardando...' : 'Guardar registro' }}
                  </button>
                </div>
              </div>
            </div>

            <!-- IMPORTAR EXCEL -->
            <div class="modal-backdrop" *ngIf="excelModal">
              <div class="modal-card excel-modal">
                <div class="modal-header">
                  <div>
                    <h2>Importar inventario</h2>
                    <p>Carga múltiples bienes mediante una plantilla Excel.</p>
                  </div>
                  <button class="close-button" (click)="closeExcel()">×</button>
                </div>

                <div class="excel-instructions">
                  <span class="excel-icon">▤</span>
                  <div>
                    <strong>Importación de Excel</strong>
                    <p>Selecciona un archivo .xlsx o .xls para revisar su contenido.</p>
                  </div>
                </div>

                <label class="upload-zone" [class.has-file]="excelFile">
                  <input type="file" accept=".xlsx,.xls" (change)="handleExcel($event)">
                  <span class="upload-symbol">↑</span>
                  <strong>{{ excelFile ? excelFile.name : 'Arrastra tu archivo aquí' }}</strong>
                  <small>
                    {{ excelFile ? (excelFile.size / 1024 | number:'1.0-0') + ' KB' : 'Selecciona un archivo desde tu equipo' }}
                  </small>
                  <span class="btn-outline upload-button">Seleccionar archivo</span>
                </label>

                <div class="excel-error" *ngIf="excelError" style="white-space: pre-line;">
                  {{ excelError }}
                </div>

                <div *ngIf="excelImportResult" style="margin: 0 24px 15px; color: #21835a; font-size: 14px;">
                  {{ excelImportResult }}
                </div>

                <div class="preview-section" *ngIf="excelRows.length">
                  <div class="preview-heading">
                    <div>
                      <strong>Vista previa</strong>
                      <p>{{ excelRows.length }} registros leídos en la hoja seleccionada.</p>
                    </div>
                    <span class="status status-good">Archivo leído</span>
                  </div>

                  <div class="table-container preview-table">
                    <table>
                      <thead>
                        <tr>
                          <th *ngFor="let h of excelHeaders">{{ h }}</th>
                        </tr>
                      </thead>
                      <tbody>
                        <tr *ngFor="let row of excelRows.slice(0,5)">
                          <td *ngFor="let h of excelHeaders">{{ row[h] ?? '—' }}</td>
                        </tr>
                      </tbody>
                    </table>
                  </div>

                  <small class="preview-note">
                    Al confirmar, los registros se enviarán al backend para guardarse en la base de datos.
                  </small>
                </div>

                <div class="modal-footer">
                  <button class="btn-outline" (click)="closeExcel()">Cancelar</button>
                  <button class="btn-primary" [disabled]="!excelRows.length || excelImporting" (click)="confirmExcel()">
                    {{ excelImporting ? 'Importando...' : 'Confirmar archivo' }}
                  </button>
                </div>
              </div>
            </div>

          </section>

          <!-- SOLICITUDES -->
          <section *ngIf="section==='solicitudes'">
            <div class="page-heading">
              <div>
                <div class="eyebrow">CONTROL DE AUTORIZACIONES</div>
                <h1>{{ vistaSolicitudes==='autorizacion' ? 'Bandeja de autorización' : 'Mis solicitudes' }}</h1>
                <p>{{ vistaSolicitudes==='autorizacion'
                  ? 'Revisa, aprueba o rechaza las altas, modificaciones y bajas enviadas por los empleados.'
                  : 'Da seguimiento a las altas, modificaciones y bajas que enviaste para autorización.' }}</p>
              </div>
              <div class="heading-actions">
                <div class="segmented" *ngIf="canAuthorize && !isAdmin">
                  <button [class.selected]="vistaSolicitudes==='autorizacion'" (click)="cambiarVistaSolicitudes('autorizacion')">Por autorizar</button>
                  <button [class.selected]="vistaSolicitudes==='mias'" (click)="cambiarVistaSolicitudes('mias')">Mis solicitudes</button>
                </div>
                <button class="btn-outline" [disabled]="solicitudesLoading" (click)="loadSolicitudes()">
                  <span>⟳</span> {{ solicitudesLoading ? 'Actualizando...' : 'Actualizar' }}
                </button>
              </div>
            </div>

            <!-- INDICADORES -->
            <div class="stats-grid">
              <button class="stat-card stat-button" [class.stat-selected]="tabSolicitudes==='PENDIENTE'" (click)="tabSolicitudes='PENDIENTE'"
                      [class.stat-alerta]="vistaSolicitudes==='autorizacion' && conteoSolicitudes('PENDIENTE') > 0">
                <div class="stat-top"><span>Pendientes</span><span class="stat-icon orange">◷</span></div>
                <strong>{{ conteoSolicitudes('PENDIENTE') }}</strong>
                <small>{{ vistaSolicitudes==='autorizacion' ? 'Esperan tu revisión' : 'En espera de revisión' }}</small>
              </button>
              <button class="stat-card stat-button" [class.stat-selected]="tabSolicitudes==='APROBADA'" (click)="tabSolicitudes='APROBADA'">
                <div class="stat-top"><span>Aprobadas</span><span class="stat-icon green">✓</span></div>
                <strong>{{ conteoSolicitudes('APROBADA') }}</strong>
                <small>Aplicadas al inventario</small>
              </button>
              <button class="stat-card stat-button" [class.stat-selected]="tabSolicitudes==='RECHAZADA'" (click)="tabSolicitudes='RECHAZADA'">
                <div class="stat-top"><span>Rechazadas</span><span class="stat-icon red">✕</span></div>
                <strong>{{ conteoSolicitudes('RECHAZADA') }}</strong>
                <small>Con motivo de rechazo</small>
              </button>
              <button class="stat-card stat-button" [class.stat-selected]="tabSolicitudes==='TODAS'" (click)="tabSolicitudes='TODAS'">
                <div class="stat-top"><span>Total</span><span class="stat-icon blue">▣</span></div>
                <strong>{{ solicitudes.length }}</strong>
                <small>Solicitudes registradas</small>
              </button>
            </div>

            <div class="inventory-panel">
              <div class="request-tabs">
                <button *ngFor="let t of tabsSolicitudes" [class.selected]="tabSolicitudes===t.key" (click)="tabSolicitudes=t.key">
                  {{ t.label }}
                  <span class="tab-count" [class.tab-alerta]="t.key==='PENDIENTE' && vistaSolicitudes==='autorizacion' && conteoSolicitudes('PENDIENTE') > 0">{{ t.key==='TODAS' ? solicitudes.length : conteoSolicitudes(t.key) }}</span>
                </button>
              </div>

              <div class="search-row request-filters">
                <div class="search-box">
                  <span>⌕</span>
                  <input [(ngModel)]="filtroSolicitudes" placeholder="Buscar por folio, clave, bien, solicitante o sede...">
                </div>
                <select [(ngModel)]="filtroOperacion" class="op-filter">
                  <option value="">Todas las operaciones</option>
                  <option value="ALTA">Alta</option>
                  <option value="MODIFICACION">Modificación</option>
                  <option value="BAJA">Baja</option>
                </select>
                <span class="results-count">{{ solicitudesFiltradas.length }} resultados</span>
              </div>

              <div class="inline-error" *ngIf="solicitudesError">{{ solicitudesError }}</div>

              <div class="table-container">
                <table>
                  <thead>
                    <tr>
                      <th>FOLIO</th>
                      <th>OPERACIÓN</th>
                      <th>BIEN</th>
                      <th *ngIf="vistaSolicitudes==='autorizacion'">SOLICITANTE</th>
                      <th>SEDE</th>
                      <th>FECHA DE SOLICITUD</th>
                      <th>ESTADO</th>
                      <th *ngIf="vistaSolicitudes==='mias'">RESOLUCIÓN</th>
                      <th></th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr *ngFor="let s of solicitudesFiltradas" class="clickable-row" (click)="abrirSolicitud(s)"
                        [class.row-pendiente]="vistaSolicitudes==='autorizacion' && s.estado==='PENDIENTE' && puedeResolver(s)">
                      <td><strong class="asset-key">#{{ s.idSolicitud }}</strong></td>
                      <td><span class="op-badge" [ngClass]="'op-' + (s.tipoOperacion || '').toLowerCase()">{{ etiquetaOperacion(s.tipoOperacion) }}</span></td>
                      <td>
                        <div class="asset-name">
                          <strong>{{ nombreBienSolicitud(s) }}</strong>
                          <small>{{ claveBienSolicitud(s) }}</small>
                        </div>
                      </td>
                      <td *ngIf="vistaSolicitudes==='autorizacion'">
                        <div class="person-cell">
                          <span class="mini-avatar">{{ iniciales(s.nombreSolicitante || s.correoSolicitante) }}</span>
                          <div class="asset-name">
                            <strong>{{ s.nombreSolicitante || 'Usuario #' + s.idUsuarioSolicitante }}</strong>
                            <small *ngIf="s.correoSolicitante && s.correoSolicitante !== s.nombreSolicitante">{{ s.correoSolicitante }}</small>
                          </div>
                        </div>
                      </td>
                      <td>{{ s.nombreSede || '—' }}</td>
                      <td>{{ s.fechaSolicitud | date:'dd/MM/yyyy HH:mm' }}</td>
                      <td>
                        <span class="status" [ngClass]="claseEstadoSolicitud(s.estado)">{{ etiquetaEstado(s.estado) }}</span>
                        <span class="revoked-tag" *ngIf="s.estado==='PENDIENTE' && s.motivoRevocacion" [title]="'Revocada: ' + s.motivoRevocacion">↺ Revocada</span>
                      </td>
                      <td *ngIf="vistaSolicitudes==='mias'" class="resolution-cell">
                        <span *ngIf="s.estado==='PENDIENTE'" class="muted">En revisión</span>
                        <span *ngIf="s.estado!=='PENDIENTE'" [title]="s.motivoResolucion || ''">
                          {{ s.motivoResolucion || (s.estado==='APROBADA' ? 'Aprobada sin comentarios' : '—') }}
                        </span>
                      </td>
                      <td class="action-cell">
                        <button class="btn-outline btn-small" [class.btn-revisar]="vistaSolicitudes==='autorizacion' && s.estado==='PENDIENTE' && puedeResolver(s)"
                                (click)="$event.stopPropagation(); abrirSolicitud(s)">
                          {{ vistaSolicitudes==='autorizacion' && s.estado==='PENDIENTE' && puedeResolver(s) ? 'Revisar ahora' : 'Ver detalle' }}
                        </button>
                      </td>
                    </tr>

                    <tr *ngIf="!solicitudesFiltradas.length && !solicitudesLoading">
                      <td colspan="9" class="empty-state">
                        <div class="empty-icon">{{ tabSolicitudes==='PENDIENTE' && vistaSolicitudes==='autorizacion' ? '✓' : '▣' }}</div>
                        <strong>{{ mensajeVacioSolicitudes }}</strong>
                        <p>{{ vistaSolicitudes==='autorizacion'
                          ? 'Las solicitudes que envíen los empleados aparecerán aquí.'
                          : 'Cuando registres, modifiques o des de baja un bien, la solicitud aparecerá aquí.' }}</p>
                      </td>
                    </tr>
                    <tr *ngIf="solicitudesLoading && !solicitudes.length">
                      <td colspan="9" class="empty-state">Cargando solicitudes...</td>
                    </tr>
                  </tbody>
                </table>
              </div>

              <div class="table-footer">
                <span>Mostrando {{ solicitudesFiltradas.length }} de {{ solicitudes.length }} solicitudes</span>
                <span>{{ vistaSolicitudes==='autorizacion' ? 'Bandeja de autorización' : 'Mis solicitudes' }}</span>
              </div>
            </div>

            <!-- DETALLE DE SOLICITUD -->
            <div class="modal-backdrop" *ngIf="solicitudSeleccionada as s" (click)="cerrarSolicitud()">
              <div class="modal-card large-modal" (click)="$event.stopPropagation()">
                <div class="modal-header">
                  <div>
                    <div class="modal-title-row">
                      <h2>Solicitud #{{ s.idSolicitud }}</h2>
                      <span class="status" [ngClass]="claseEstadoSolicitud(s.estado)">{{ etiquetaEstado(s.estado) }}</span>
                    </div>
                    <p>{{ descripcionOperacion(s) }}</p>
                  </div>
                  <button class="close-button" [disabled]="resolviendo" (click)="cerrarSolicitud()">×</button>
                </div>

                <div class="request-summary">
                  <div><small>Operación</small><span class="op-badge" [ngClass]="'op-' + (s.tipoOperacion || '').toLowerCase()">{{ etiquetaOperacion(s.tipoOperacion) }}</span></div>
                  <div><small>Solicitante</small><strong>{{ s.nombreSolicitante || 'Usuario #' + s.idUsuarioSolicitante }}</strong></div>
                  <div><small>Sede</small><strong>{{ s.nombreSede || '—' }}</strong></div>
                  <div><small>Fecha de solicitud</small><strong>{{ s.fechaSolicitud | date:'dd/MM/yyyy HH:mm' }}</strong></div>
                </div>

                <div class="form-section" *ngIf="s.motivoSolicitud">
                  <h3>Motivo de la solicitud</h3>
                  <div class="note-box">{{ s.motivoSolicitud }}</div>
                </div>

                <!-- BAJA -->
                <div class="form-section" *ngIf="s.tipoOperacion==='BAJA'">
                  <div class="warning-box">
                    <strong>Baja de bien</strong>
                    <p>Al aprobarse, el bien <b>{{ nombreBienSolicitud(s) }}</b> ({{ claveBienSolicitud(s) }}) se retirará del inventario activo.</p>
                  </div>
                </div>

                <!-- ALTA / MODIFICACIÓN -->
                <div class="form-section" *ngIf="s.tipoOperacion!=='BAJA'">
                  <h3>{{ s.tipoOperacion==='ALTA' ? 'Datos del nuevo bien' : 'Cambios propuestos' }}</h3>
                  <p class="section-hint" *ngIf="s.tipoOperacion==='MODIFICACION' && hayValoresActuales">
                    Los campos resaltados son los que cambian respecto al registro actual.
                  </p>
                  <div class="table-container compare-table">
                    <table>
                      <thead>
                        <tr>
                          <th>CAMPO</th>
                          <th *ngIf="s.tipoOperacion==='MODIFICACION' && hayValoresActuales">VALOR ACTUAL</th>
                          <th>{{ s.tipoOperacion==='ALTA' ? 'VALOR' : 'VALOR PROPUESTO' }}</th>
                        </tr>
                      </thead>
                      <tbody>
                        <tr *ngFor="let c of camposSolicitud" [class.changed-row]="c.cambio">
                          <td class="field-label">{{ c.etiqueta }}</td>
                          <td *ngIf="s.tipoOperacion==='MODIFICACION' && hayValoresActuales" class="old-value">{{ c.actual }}</td>
                          <td class="new-value">{{ c.propuesto }}</td>
                        </tr>
                        <tr *ngIf="!camposSolicitud.length">
                          <td colspan="3" class="empty-state">La solicitud no contiene datos.</td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                </div>

                <!-- RESOLUCIÓN YA EMITIDA -->
                <div class="form-section" *ngIf="s.estado!=='PENDIENTE'">
                  <div class="resolution-box" [class.approved]="s.estado==='APROBADA'" [class.rejected]="s.estado==='RECHAZADA'">
                    <strong>{{ s.estado==='APROBADA' ? '✓ Solicitud aprobada' : '✕ Solicitud rechazada' }}</strong>
                    <p>
                      {{ s.nombreRevisor ? 'Revisó: ' + s.nombreRevisor : '' }}
                      {{ s.fechaResolucion ? ' · ' + (s.fechaResolucion | date:'dd/MM/yyyy HH:mm') : '' }}
                    </p>
                    <div class="resolution-reason" *ngIf="s.motivoResolucion">
                      <small>{{ s.estado==='RECHAZADA' ? 'Motivo del rechazo' : 'Comentario' }}</small>
                      {{ s.motivoResolucion }}
                    </div>
                  </div>
                </div>

                <!-- REVOCACIÓN: ACCIÓN DEL ADMINISTRADOR -->
                <div class="form-section" *ngIf="isAdmin && s.estado!=='PENDIENTE'">
                  <ng-container *ngIf="s.revocable; else noRevocable">
                    <button class="revoke-link" *ngIf="!modoRevocar" (click)="iniciarRevocacion()">
                      ↺ ¿Te equivocaste? Revocar esta autorización
                    </button>
                    <div class="revoke-box" *ngIf="modoRevocar">
                      <strong>↺ Revocar autorización</strong>
                      <ul>
                        <li *ngIf="s.estado==='APROBADA' && s.tipoOperacion==='ALTA'">Se <b>eliminará el bien</b> que se creó con esta solicitud.</li>
                        <li *ngIf="s.estado==='APROBADA' && s.tipoOperacion==='MODIFICACION'">El bien <b>regresará a los datos que tenía</b> antes de este cambio.</li>
                        <li *ngIf="s.estado==='APROBADA' && s.tipoOperacion==='BAJA'">El bien <b>volverá al inventario activo</b>.</li>
                        <li *ngIf="s.estado==='RECHAZADA'">Se <b>anulará el rechazo</b> y su motivo.</li>
                        <li>La solicitud regresará a la <b>bandeja de autorización como pendiente</b> para resolverla de nuevo.</li>
                      </ul>
                      <label>Motivo de la revocación <span class="muted">(obligatorio)</span></label>
                      <textarea rows="2" maxlength="500" [(ngModel)]="motivoRevocacion" [class.input-error]="revocacionRequerida"
                                placeholder="Ej. Se aprobó por error, faltaba validar el número de serie." (input)="revocacionRequerida=false"></textarea>
                      <div class="field-footer">
                        <span class="field-error" *ngIf="revocacionRequerida">Escribe por qué revocas la autorización.</span>
                        <span class="char-count">{{ (motivoRevocacion || '').length }}/500</span>
                      </div>
                      <div class="inline-error" *ngIf="errorRevocacion">{{ errorRevocacion }}</div>
                    </div>
                  </ng-container>
                  <ng-template #noRevocable>
                    <p class="revoke-off" *ngIf="s.estado==='APROBADA'">Esta autorización no se puede revocar porque se aprobó antes de habilitar esta función (no hay copia de los datos anteriores).</p>
                  </ng-template>
                </div>

                <!-- PENDIENTE QUE YA HABÍA SIDO RESUELTA Y SE REVOCÓ -->
                <div class="form-section" *ngIf="s.estado==='PENDIENTE' && s.motivoRevocacion">
                  <div class="revoked-note">
                    <strong>↺ Autorización revocada</strong>
                    <p>{{ s.nombreRevoco ? s.nombreRevoco + ' · ' : '' }}{{ s.fechaRevocacion | date:'dd/MM/yyyy HH:mm' }}</p>
                    <div>{{ s.motivoRevocacion }}</div>
                  </div>
                </div>

                <!-- PENDIENTE: VISTA DEL EMPLEADO -->
                <div class="form-section" *ngIf="s.estado==='PENDIENTE' && !puedeResolver(s)">
                  <div class="info-box">
                    <strong>◷ En espera de autorización</strong>
                    <p>Un administrador revisará esta solicitud. Los cambios se aplicarán al inventario solo si se aprueba.</p>
                  </div>
                </div>

                <!-- PENDIENTE: ACCIONES DEL ADMINISTRADOR -->
                <div class="form-section decision-section" *ngIf="s.estado==='PENDIENTE' && puedeResolver(s)">
                  <h3>Resolución</h3>
                  <label>Comentario <span class="muted">(obligatorio para rechazar)</span></label>
                  <textarea rows="3" maxlength="500" [(ngModel)]="motivoResolucion" [class.input-error]="motivoRequerido"
                            placeholder="Escribe un comentario o el motivo del rechazo..." (input)="motivoRequerido=false"></textarea>
                  <div class="field-footer">
                    <span class="field-error" *ngIf="motivoRequerido">Indica el motivo del rechazo para que el solicitante sepa qué corregir.</span>
                    <span class="char-count">{{ (motivoResolucion || '').length }}/500</span>
                  </div>
                  <div class="inline-error" *ngIf="errorResolucion">{{ errorResolucion }}</div>
                </div>

                <div class="modal-footer">
                  <ng-container *ngIf="s.estado==='PENDIENTE' && puedeResolver(s); else soloCerrar">
                    <button class="btn-outline" [disabled]="resolviendo" (click)="cerrarSolicitud()">Cancelar</button>
                    <button class="btn-danger" [disabled]="resolviendo" (click)="resolverSolicitud(false)">
                      {{ resolviendo==='rechazar' ? 'Rechazando...' : 'Rechazar' }}
                    </button>
                    <button class="btn-success" [disabled]="resolviendo" (click)="resolverSolicitud(true)">
                      {{ resolviendo==='aprobar' ? 'Aprobando...' : 'Aprobar y aplicar' }}
                    </button>
                  </ng-container>
                  <ng-template #soloCerrar>
                    <ng-container *ngIf="modoRevocar; else botonCerrar">
                      <button class="btn-outline" [disabled]="revocando" (click)="modoRevocar=false">Cancelar</button>
                      <button class="btn-revoke" [disabled]="revocando" (click)="revocarSolicitud()">
                        <span class="mini-ring light" *ngIf="revocando"></span>
                        {{ revocando ? 'Revocando…' : 'Revocar y regresar a pendientes' }}
                      </button>
                    </ng-container>
                    <ng-template #botonCerrar>
                      <button class="btn-outline" *ngIf="s.estado==='APROBADA' && (s.tipoOperacion==='ALTA' || s.tipoOperacion==='BAJA')"
                              (click)="formatoDeSolicitud(s)">
                        📄 Formato de {{ s.tipoOperacion==='ALTA' ? 'alta' : 'baja' }}
                      </button>
                      <button class="btn-primary" (click)="cerrarSolicitud()">Cerrar</button>
                    </ng-template>
                  </ng-template>
                </div>
              </div>
            </div>
          </section>

          <!-- INVENTARIO ANUAL -->
          <section *ngIf="section==='anual'">

            <!-- ===== Lista de levantamientos ===== -->
            <ng-container *ngIf="anualVista==='lista'">
              <div class="inv-head">
                <div>
                  <h1>Levantamiento de inventario</h1>
                  <p class="inv-resumen">Por sede y por persona: escanea cada bien y se arman solos los resguardos y el acta de cierre.</p>
                </div>
                <div class="heading-actions">
                  <button class="btn-outline" [disabled]="levCargando" (click)="loadLevantamientos()">⟳ Actualizar</button>
                  <button class="btn-primary" *ngIf="isAdmin" (click)="abrirNuevoLevantamiento()">＋ Nuevo levantamiento</button>
                </div>
              </div>

              <div class="inventory-panel placeholder-panel" *ngIf="!levCargando && !levantamientos.length">
                <div class="placeholder-icon">✓</div>
                <h2>Aún no hay levantamientos</h2>
                <p>{{ isAdmin ? 'Crea el primero: se genera el Acta de inicio y el capturista podrá revisar los bienes desde su celular.' : 'Cuando el administrador te asigne como capturista, aparecerá aquí.' }}</p>
                <button class="btn-primary" *ngIf="isAdmin" (click)="abrirNuevoLevantamiento()">＋ Nuevo levantamiento</button>
              </div>

              <div class="lev-list">
                <div class="lev-card" *ngFor="let l of levantamientos" [class.open]="levAbierto===l.idLevantamiento" [class.cerrado]="l.estado==='CERRADO'">
                  <button class="lev-head" (click)="levAbierto = levAbierto===l.idLevantamiento ? null : l.idLevantamiento">
                    <span class="lev-ico">{{ l.estado==='CERRADO' ? '✓' : '◷' }}</span>
                    <span class="lev-main">
                      <strong>Inventario {{ l.anio }} · {{ l.sede }}</strong>
                      <span class="lev-bar"><span [style.width.%]="pctLev(l)"></span></span>
                      <small>{{ l.capturados | number }} de {{ l.totalBienes | number }} bienes revisados · {{ pctLev(l) }}%</small>
                    </span>
                    <span class="lev-chip" [class.ok]="l.estado==='CERRADO'">{{ l.estado==='CERRADO' ? 'Cerrado' : 'En curso' }}</span>
                    <span class="lev-chev">⌄</span>
                  </button>

                  <div class="lev-body" *ngIf="levAbierto===l.idLevantamiento">
                    <div class="lev-stats">
                      <div><b>{{ l.capturados | number }}</b><small>Revisados</small></div>
                      <div><b>{{ pendientesLev(l) | number }}</b><small>{{ l.estado==='CERRADO' ? 'No localizados' : 'Pendientes' }}</small></div>
                      <div><b>{{ (l.noEncontrados || 0) | number }}</b><small>Marcados no encontrados</small></div>
                      <div><b>{{ ((l.nuevos || 0) + (l.transferidos || 0)) | number }}</b><small>Nuevos o traídos</small></div>
                    </div>

                    <div class="lev-actions">
                      <button class="btn-primary" *ngIf="l.estado==='EN_CURSO' && (l.soyCapturista || l.puedoAdministrar)" (click)="iniciarCaptura(l)">📷 Capturar bienes</button>
                      <button class="btn-outline" (click)="abrirAvance(l)">📊 Ver avance</button>
                    </div>

                    <div class="lev-docs">
                      <span class="lev-docs-label">Documentos</span>
                      <button *ngIf="l.idDocumentoInicio" (click)="abrirDocumentoGuardado(l.idDocumentoInicio)">📄 Acta de inicio</button>
                      <button *ngIf="l.idDocumentoCierre" (click)="abrirDocumentoGuardado(l.idDocumentoCierre)">📄 Acta de cierre <em *ngIf="l.estado==='EN_CURSO'">borrador</em></button>
                      <button *ngIf="l.resguardos" (click)="verResguardosLev(l)">👤 Resguardos <em>{{ l.resguardos }}</em></button>
                    </div>

                    <div class="lev-actions" *ngIf="l.puedoAdministrar">
                      <button class="btn-danger" *ngIf="l.estado==='EN_CURSO'" (click)="cerrarLevantamiento(l)">Cerrar levantamiento</button>
                      <button class="btn-eliminar-lev" (click)="eliminarLevantamiento(l)">🗑 Eliminar inventario</button>
                    </div>

                    <button class="lev-more" (click)="levDetalle = levDetalle===l.idLevantamiento ? null : l.idLevantamiento">
                      {{ levDetalle===l.idLevantamiento ? '▴ Ocultar detalles' : '▾ Ver detalles' }}
                    </button>
                    <div class="lev-detail" *ngIf="levDetalle===l.idLevantamiento">
                      <dl>
                        <div><dt>Inicio</dt><dd>{{ l.fechaInicio | date:'dd/MM/yyyy HH:mm' }}</dd></div>
                        <div *ngIf="l.fechaCierre"><dt>Cierre</dt><dd>{{ l.fechaCierre | date:'dd/MM/yyyy HH:mm' }}</dd></div>
                        <div><dt>Captura</dt><dd>{{ l.capturista || '—' }}</dd></div>
                        <div><dt>Lugar</dt><dd>{{ l.ciudad || '—' }}<br><small>{{ l.direccion }}</small></dd></div>
                        <div><dt>Pregunta</dt><dd>{{ l.pregunta?.texto }}<br><small>{{ (l.pregunta?.opciones || []).join(' · ') }}</small></dd></div>
                        <div><dt>Participantes</dt><dd><span class="lev-person" *ngFor="let p of l.participantes">{{ p.nombre }} <small>{{ p.cargo }}</small></span></dd></div>
                        <div><dt>Revisión</dt><dd>{{ l.fueraDeLugar | number }} fuera de lugar · {{ l.fotosNuevas | number }} fotos tomadas<br><small>Fotos en {{ l.almacenFotos }}</small></dd></div>
                      </dl>
                    </div>
                  </div>
                </div>
              </div>
            </ng-container>

            <!-- ===== Avance (monitoreo): resumen -> areas -> bienes ===== -->
            <div class="cap av" *ngIf="av">
              <div class="cap-top">
                <button class="doc-back" (click)="cerrarAvance()">{{ av.enCaptura ? '← Volver a capturar' : '← Inventarios' }}</button>
                <strong>Avance · Inventario {{ av.lev.anio }} · {{ av.lev.sede }}</strong>
                <span class="lev-bar"><span [style.width.%]="av.vista.pct"></span></span>
                <small>{{ av.vista.revisados | number }} de {{ av.vista.total | number }} revisados · {{ av.vista.pct }}%
                  <ng-container *ngIf="!av.enCaptura && av.actualizado"> · {{ av.lev.estado==='EN_CURSO' ? 'se actualiza solo · ' : '' }}{{ av.actualizado | date:'HH:mm' }}</ng-container>
                </small>
              </div>

              <div class="cap-loading" *ngIf="av.cargando"><div class="loader-ring"></div><strong>Cargando avance…</strong></div>
              <div class="inline-error" *ngIf="av.error">{{ av.error }}</div>

              <ng-container *ngIf="!av.cargando">
                <div class="av-aviso etq-aviso" *ngIf="av.enCaptura && cap?.etiquetaPorAsignar">
                  <b>🏷 Etiqueta nueva {{ cap.etiquetaPorAsignar }}</b>
                  <span>Toca el bien al que se la pegaste.</span>
                  <button class="btn-outline btn-small" (click)="cap.etiquetaPorAsignar=null; cerrarAvance()">Cancelar</button>
                </div>
                <div class="av-aviso" *ngIf="av.antesDeCerrar">
                  <b>Antes de cerrar: faltan {{ av.vista.cuentas.pendientes | number }} bienes por revisar.</b>
                  <span>Revísalos por persona o por área. Si cierras así, quedarán como "no localizados" en el acta.</span>
                  <button class="btn-danger" (click)="cerrarDesdeAvance()">Cerrar de todos modos</button>
                </div>

                <div class="seg">
                  <button [class.on]="av.agrupar==='persona'" (click)="avAgrupar('persona')">👤 Por persona</button>
                  <button [class.on]="av.agrupar==='area'" (click)="avAgrupar('area')">📍 Por área</button>
                </div>

                <!-- Nivel 1: resumen (cada numero es un filtro) -->
                <div class="av-chips">
                  <button *ngFor="let f of avFiltros" [class.on]="av.filtro===f.key" [ngClass]="'f-' + f.key" (click)="avFiltrar(f.key)">
                    <b>{{ av.vista.cuentas[f.key] | number }}</b><small>{{ f.ico }} {{ f.label }}</small>
                  </button>
                </div>

                <input class="cap-search" [(ngModel)]="av.busqueda" (ngModelChange)="avRecalcular()" placeholder="Buscar por clave, artículo o serie...">

                <!-- Con busqueda: resultados directos -->
                <div class="av-items av-flat" *ngIf="av.busqueda?.trim()">
                  <ng-container *ngFor="let b of av.vista.resultados">
                    <ng-container *ngTemplateOutlet="avItemTpl; context: { $implicit: b, conArea: true }"></ng-container>
                  </ng-container>
                  <p class="cap-hint" *ngIf="!av.vista.resultados.length">Ningún bien coincide con la búsqueda.</p>
                </div>

                <!-- Nivel 2: areas (cerradas) -->
                <div class="av-areas" *ngIf="!av.busqueda?.trim()">
                  <div class="av-area" *ngFor="let a of av.vista.areas; trackBy: trackArea" [class.open]="av.areaAbierta===a.key">
                    <button class="av-area-head" (click)="av.areaAbierta = av.areaAbierta===a.key ? null : a.key">
                      <span class="av-area-main">
                        <strong>{{ a.nombre }}</strong>
                        <span class="lev-bar"><span [style.width.%]="a.pct"></span></span>
                      </span>
                      <span class="av-area-count">{{ a.revisados }} / {{ a.total }}</span>
                      <span class="av-tag ok" *ngIf="a.revisados===a.total">✓ Completa</span>
                      <span class="av-tag" *ngIf="!a.revisados && !a.no">Sin iniciar</span>
                      <span class="av-tag warn" *ngIf="a.no">✗ {{ a.no }}</span>
                      <span class="av-tag warn" *ngIf="a.fuera">⚠ {{ a.fuera }}</span>
                      <span class="lev-chev">⌄</span>
                    </button>
                    <!-- Nivel 3: bienes del area -->
                    <div class="av-items" *ngIf="av.areaAbierta===a.key">
                      <ng-container *ngFor="let b of a.items.slice(0, avLimite(a.key))">
                        <ng-container *ngTemplateOutlet="avItemTpl; context: { $implicit: b, conArea: false }"></ng-container>
                      </ng-container>
                      <button class="lev-more" *ngIf="a.items.length > avLimite(a.key)" (click)="av.limites[a.key] = avLimite(a.key) + 50">
                        ▾ Ver {{ a.items.length - avLimite(a.key) }} más
                      </button>
                      <button class="lev-more av-resg" *ngIf="av.agrupar==='persona' && a.key!=='sin' && a.revisados" (click)="abrirResguardo(av.lev, a.key)">
                        📄 Resguardo de {{ a.nombre }}{{ av.lev.estado==='EN_CURSO' ? ' (borrador)' : '' }}
                      </button>
                    </div>
                  </div>
                  <p class="cap-hint" *ngIf="!av.vista.areas.length">No hay bienes en este filtro. 🎉</p>
                </div>

                <button class="lev-more av-pdf" *ngIf="av.vista.cuentas.pendientes" (click)="avPdfPendientes()">🖨 Descargar lista de pendientes por {{ av.agrupar==='persona' ? 'persona' : 'área' }} (PDF)</button>
              </ng-container>

              <ng-template #avItemTpl let-b let-conArea="conArea">
                <div class="av-item" [ngClass]="'st-' + avEstado(b)">
                  <button class="av-item-row" (click)="avTocarBien(b)">
                    <span class="av-ico">{{ avEstado(b)==='ok' ? '✓' : avEstado(b)==='fuera' ? '⚠' : avEstado(b)==='no' ? '✗' : '⏳' }}</span>
                    <span class="av-txt">
                      <span><b class="chip-key">{{ b.clave }}</b> {{ b.articulo }} <em class="av-orig" *ngIf="b.origen">{{ b.origen==='NUEVO' ? 'Nuevo' : 'De ' + (b.sede_anterior || 'otra sede') }}</em></span>
                      <small>
                        <ng-container *ngIf="conArea || av.agrupar==='persona'">📍 {{ b.ubicacion || 'Sin ubicación' }} · </ng-container>
                        <ng-container *ngIf="avEstado(b)==='no'">No encontrado</ng-container>
                        <ng-container *ngIf="avEstado(b)==='ok'">{{ b.respuesta }}</ng-container>
                        <ng-container *ngIf="avEstado(b)==='fuera'">{{ b.respuesta }} · encontrado en {{ b.ubicacion_encontrada || 'otra área' }}</ng-container>
                        <ng-container *ngIf="avEstado(b)==='pendiente'">Pendiente{{ b.tiene_foto ? '' : ' · 📷 sin foto' }}</ng-container>
                      </small>
                    </span>
                    <span class="av-go" *ngIf="av.enCaptura">›</span>
                  </button>
                  <div class="av-detalle" *ngIf="!av.enCaptura && av.detalle===b.id_bien">
                    <img *ngIf="b.foto_mini" [src]="b.foto_mini" alt="Foto del bien" loading="lazy" (click)="ampliarFoto(b)">
                    <dl>
                      <div *ngIf="esTexto(b.marca) || esTexto(b.modelo)"><dt>Marca / modelo</dt><dd>{{ [b.marca, b.modelo].filter(esTexto).join(' · ') }}</dd></div>
                      <div *ngIf="esTexto(b.numero_serie)"><dt>Serie</dt><dd>{{ b.numero_serie }}</dd></div>
                      <div><dt>Persona</dt><dd>{{ b.persona || 'Sin resguardante' }}</dd></div>
                      <div><dt>Responsable</dt><dd>{{ b.responsable || 'Sin asignar' }}</dd></div>
                      <div *ngIf="b.fecha_registro"><dt>Revisado</dt><dd>{{ b.fecha_registro | date:'dd/MM/yyyy HH:mm' }}</dd></div>
                    </dl>
                  </div>
                </div>
              </ng-template>
            </div>

            <!-- ===== Captura (celular) ===== -->
            <div class="cap" *ngIf="anualVista==='captura' && cap && !av">
              <div class="cap-top">
                <button class="doc-back" (click)="salirCaptura()">← Salir</button>
                <strong>Inventario {{ cap.lev.anio }} · {{ cap.lev.sede }}</strong>
                <span class="lev-bar"><span [style.width.%]="capPct"></span></span>
                <small>{{ capRevisados | number }} de {{ cap.bienes.length | number }} revisados</small>
                <button class="cap-avance-btn" *ngIf="!cap.cargando && cap.paso!=='bien'" (click)="abrirAvanceCaptura()">📋 Ver avance</button>
              </div>

              <div class="cap-loading" *ngIf="cap.cargando"><div class="loader-ring"></div><strong>Cargando bienes de la sede…</strong></div>

              <!-- Paso 1: persona o área -->
              <ng-container *ngIf="!cap.cargando && cap.paso==='area'">
                <h2 class="cap-q">{{ cap.modo==='persona' ? '¿De quién son los bienes que vas a revisar?' : '¿En qué área estás?' }}</h2>
                <div class="seg">
                  <button [class.on]="cap.modo==='persona'" (click)="cap.modo='persona'; cap.busquedaArea=''">👤 Por persona</button>
                  <button [class.on]="cap.modo==='area'" (click)="cap.modo='area'; cap.busquedaArea=''">📍 Por área</button>
                </div>
                <p class="cap-hint" *ngIf="cap.modo==='area'">Así sabremos si algún bien está fuera de su lugar.</p>
                <input class="cap-search" [(ngModel)]="cap.busquedaArea" [placeholder]="cap.modo==='persona' ? 'Buscar persona...' : 'Buscar área...'">
                <div class="cap-areas">
                  <button *ngFor="let g of gruposCaptura" [class.done]="g.hechos===g.total" (click)="elegirGrupo(g)">
                    <span>{{ g.nombre }}</span>
                    <small>{{ g.hechos }} / {{ g.total }}</small>
                  </button>
                  <p class="cap-hint" *ngIf="!gruposCaptura.length">Sin resultados.</p>
                </div>
                <button class="lev-more" (click)="elegirGrupo(null)">Ver todos los bienes de la sede</button>
              </ng-container>

              <!-- Paso 2: escanear -->
              <ng-container *ngIf="!cap.cargando && cap.paso==='escanear'">
                <button class="cap-area-chip" (click)="cap.paso='area'">{{ cap.modo==='persona' ? '👤' : '📍' }} {{ grupoActualNombre }} · cambiar</button>
                <div class="cap-ok" *ngIf="cap.ultimo">✓ {{ cap.ultimo }}</div>
                <div class="inline-error" *ngIf="cap.error">{{ cap.error }}</div>
                <button class="lev-more" *ngIf="cap.ofrecerNuevo" (click)="abrirNuevoBienCap()">＋ Registrarlo como bien nuevo</button>

                <!-- Recién agregado (nuevo o traído): etiqueta y seguir -->
                <div class="cap-agregado" *ngIf="cap.agregado as ag">
                  <b>{{ ag.tipo==='nuevo' ? '✓ Bien nuevo registrado' : '✓ Bien traído a ' + cap.lev.sede }}</b>
                  <span><b class="chip-key">{{ ag.bien.clave }}</b> {{ ag.bien.articulo }}</span>
                  <small>{{ ag.bien.etiqueta ? 'Etiqueta ' + ag.bien.etiqueta + ' asignada.' : 'Pégale una etiqueta en blanco y escanéala: se le asigna sola.' }}</small>
                  <div class="cap-agregado-btns">
                    <button class="btn-primary" (click)="abrirBienCaptura(ag.bien)">Continuar con foto y estado →</button>
                  </div>
                </div>

                <div class="cap-scanner" *ngIf="cap.escaneando">
                  <video #camQr playsinline muted autoplay (click)="camQr.play()"></video>
                  <span class="cap-frame"></span>
                  <button class="btn-outline" (click)="detenerEscaner()">Cancelar</button>
                </div>

                <ng-container *ngIf="!cap.escaneando">
                  <button class="cap-big" *ngIf="camaraEnVivo" (click)="iniciarEscaner()">
                    <span>📷</span> Escanear QR
                  </button>
                  <label class="cap-big" [class.secondary]="camaraEnVivo">
                    <span>🖼</span> {{ camaraEnVivo ? 'Leer QR desde una foto' : 'Tomar foto del QR' }}
                    <input type="file" accept="image/*" capture="environment" (change)="leerQrDeFoto($event)" hidden>
                  </label>
                  <p class="cap-hint" *ngIf="!camaraEnVivo">Para escanear en vivo, el celular necesita abrir DOZZER con https. Mientras, toma una foto del QR.</p>

                  <button class="lev-more" (click)="cap.verClave = !cap.verClave">{{ cap.verClave ? '▴' : '▾' }} Escribir la clave a mano</button>
                  <div class="cap-clave" *ngIf="cap.verClave">
                    <input [(ngModel)]="cap.clave" placeholder="Ej. CE0004" (keyup.enter)="buscarPorClave()">
                    <button class="btn-primary" (click)="buscarPorClave()">Buscar</button>
                  </div>

                  <div class="cap-tools">
                    <button (click)="abrirAvanceCaptura('pendientes')"><span>📋</span><b>{{ pendientesGrupo.length | number }}</b><small>Pendientes</small></button>
                    <button (click)="abrirAvanceCaptura('sinetiq')"><span>🏷</span><b>{{ sinEtiquetaGrupo | number }}</b><small>Sin etiqueta</small></button>
                    <button (click)="abrirNuevoBienCap()"><span>＋</span><b>Nuevo</b><small>Bien que no está</small></button>
                    <button (click)="abrirBuscarGeneral()"><span>🔍</span><b>Buscar</b><small>En todas las sedes</small></button>
                  </div>
                </ng-container>
              </ng-container>

              <!-- Paso 3: el bien -->
              <ng-container *ngIf="!cap.cargando && cap.paso==='bien' && cap.bien as b">
                <div class="cap-bien">
                  <b class="chip-key">{{ b.clave }}</b>
                  <h2>{{ b.articulo }}</h2>
                  <small>{{ [b.marca, b.modelo].filter(esTexto).join(' · ') }}</small>
                  <small>📍 Registrado en: {{ b.ubicacion || 'Sin ubicación' }}</small>
                  <small>👤 {{ b.persona || 'Sin resguardante' }}</small>
                  <small>🏷 {{ b.etiqueta || 'Sin etiqueta: escanea una en blanco para asignarla' }}</small>
                  <div class="cap-warn" *ngIf="cap.area && b.id_ubicacion !== cap.area.id">⚠ Está registrado en otra área: se marcará como fuera de lugar.</div>
                  <div class="cap-info" *ngIf="b.estatus==='REVISADO'">Ya se revisó antes ({{ b.respuesta }}). Si guardas, se actualiza.</div>
                  <div class="cap-info" *ngIf="b.estatus==='NO_ENCONTRADO'">Estaba marcado como no encontrado. Si lo guardas, queda revisado.</div>
                </div>

                <div class="cap-step" *ngIf="!b.tiene_foto">
                  <div class="step-title"><span>1</span> Foto del bien</div>
                  <label class="cap-photo" *ngIf="!cap.fotoUrl">
                    <span>📷</span> Tomar foto
                    <input type="file" accept="image/*" capture="environment" (change)="tomarFotoBien($event)" hidden>
                  </label>
                  <div class="cap-preview" *ngIf="cap.fotoUrl">
                    <img [src]="cap.fotoUrl" alt="Foto del bien">
                    <label class="btn-outline btn-small">Repetir<input type="file" accept="image/*" capture="environment" (change)="tomarFotoBien($event)" hidden></label>
                  </div>
                </div>
                <div class="cap-info cap-tiene-foto" *ngIf="b.tiene_foto">
                  <img *ngIf="b.foto_mini" [src]="b.foto_mini" alt="Foto del bien" loading="lazy" (click)="ampliarFoto(b)">
                  <span>✓ Este bien ya tiene foto</span>
                </div>

                <div class="cap-step" *ngIf="b.tiene_foto || cap.fotoUrl">
                  <div class="step-title"><span>{{ b.tiene_foto ? 1 : 2 }}</span> {{ cap.lev.pregunta?.texto }}</div>
                  <div class="cap-options">
                    <button *ngFor="let o of cap.lev.pregunta?.opciones" [class.on]="cap.respuesta===o" (click)="cap.respuesta=o">{{ o }}</button>
                  </div>
                </div>

                <div class="inline-error" *ngIf="cap.error">{{ cap.error }}</div>
                <div class="cap-footer">
                  <button class="btn-outline" [disabled]="cap.guardando" (click)="cancelarBienCaptura()">Cancelar</button>
                  <button class="btn-primary" [disabled]="cap.guardando || !cap.respuesta || (!b.tiene_foto && !cap.foto)" (click)="guardarCaptura()">
                    <span class="mini-ring light" *ngIf="cap.guardando"></span>
                    {{ cap.guardando ? 'Guardando…' : 'Guardar y siguiente' }}
                  </button>
                </div>
                <button class="cap-noenc" *ngIf="cap.desdeLista && b.estatus!=='NO_ENCONTRADO'" [disabled]="cap.guardando" (click)="marcarNoEncontradoCap()">
                  ✗ No lo encuentro · marcar como no encontrado
                </button>
              </ng-container>
            </div>

            <!-- ===== Bien nuevo (durante la captura) ===== -->
            <div class="modal-backdrop" *ngIf="capNuevo">
              <div class="modal-card wizard">
                <div class="modal-header">
                  <div>
                    <h2>＋ Bien nuevo</h2>
                    <p>Entra directo al inventario de {{ cap?.lev?.sede }}, sin autorización.</p>
                  </div>
                  <button class="close-button" [disabled]="capNuevo.guardando" (click)="capNuevo=null">×</button>
                </div>
                <div class="form-section">
                  <label>¿Qué es? *</label>
                  <input [(ngModel)]="capNuevo.articulo" maxlength="200" placeholder="Ej. Silla ejecutiva">
                  <label>Área *</label>
                  <select [(ngModel)]="capNuevo.idUbicacion">
                    <option [ngValue]="null">Selecciona el área</option>
                    <option *ngFor="let a of capOpciones.areas" [ngValue]="a.id">{{ a.nombre }}</option>
                  </select>
                  <label>Persona <span class="muted">(de quién es el resguardo)</span></label>
                  <select [(ngModel)]="capNuevo.idPersona">
                    <option [ngValue]="null">Sin resguardante</option>
                    <option *ngFor="let x of capOpciones.personas" [ngValue]="x.id">{{ x.nombre }}</option>
                  </select>
                  <button class="lev-more" (click)="capNuevo.mas = !capNuevo.mas">{{ capNuevo.mas ? '▴ Menos datos' : '▾ Más datos (marca, modelo, serie…)' }}</button>
                  <ng-container *ngIf="capNuevo.mas">
                    <label>Clasificación</label>
                    <select [(ngModel)]="capNuevo.idClasificacion">
                      <option [ngValue]="null">Sin clasificación</option>
                      <option *ngFor="let c of capOpciones.clasificaciones" [ngValue]="c.id">{{ c.nombre }}</option>
                    </select>
                    <label>Marca</label>
                    <select [(ngModel)]="capNuevo.idMarca">
                      <option [ngValue]="null">Sin marca</option>
                      <option *ngFor="let m of capOpciones.marcas" [ngValue]="m.id">{{ m.nombre }}</option>
                    </select>
                    <label>Modelo</label>
                    <input [(ngModel)]="capNuevo.modelo" maxlength="150">
                    <label>Número de serie</label>
                    <input [(ngModel)]="capNuevo.numeroSerie" maxlength="150">
                    <label>Características</label>
                    <textarea rows="2" [(ngModel)]="capNuevo.caracteristicas" maxlength="500" placeholder="Color, tamaño, material…"></textarea>
                    <label>Clave <span class="muted">(vacía = se asigna sola)</span></label>
                    <input [(ngModel)]="capNuevo.clave" maxlength="50" placeholder="Ej. CE1850">
                  </ng-container>
                </div>
                <div class="inline-error" *ngIf="capNuevo.error" style="margin:0 24px 12px">{{ capNuevo.error }}</div>
                <div class="modal-footer">
                  <button class="btn-outline" [disabled]="capNuevo.guardando" (click)="capNuevo=null">Cancelar</button>
                  <button class="btn-primary" [disabled]="capNuevo.guardando" (click)="guardarNuevoBienCap()">
                    <span class="mini-ring light" *ngIf="capNuevo.guardando"></span>
                    {{ capNuevo.guardando ? 'Registrando…' : 'Registrar bien' }}
                  </button>
                </div>
              </div>
            </div>

            <!-- ===== Buscar en todo el inventario (todas las sedes) ===== -->
            <div class="modal-backdrop" *ngIf="capBuscar">
              <div class="modal-card wizard">
                <div class="modal-header">
                  <div>
                    <h2>🔍 Buscar en todo el inventario</h2>
                    <p>Para un bien que está aquí pero el sistema tiene en otra sede.</p>
                  </div>
                  <button class="close-button" [disabled]="capBuscar.guardando" (click)="capBuscar=null">×</button>
                </div>
                <div class="form-section">
                  <div class="cap-clave">
                    <input [(ngModel)]="capBuscar.q" placeholder="Clave, serie o nombre del bien" (keyup.enter)="buscarGeneralCap()">
                    <button class="btn-primary" [disabled]="capBuscar.cargando" (click)="buscarGeneralCap()">Buscar</button>
                  </div>
                  <p class="cap-hint" *ngIf="capBuscar.cargando">Buscando en todas las sedes…</p>
                  <div class="busq-list" *ngIf="!capBuscar.cargando">
                    <div class="busq-item" *ngFor="let r of capBuscar.resultados" [class.open]="capBuscar.elegido===r">
                      <div class="busq-row">
                        <span class="av-txt">
                          <span><b class="chip-key">{{ r.clave }}</b> {{ r.articulo }}</span>
                          <small>📍 {{ r.sede || 'Sin sede' }} · {{ r.ubicacion || 'Sin área' }}{{ r.persona ? ' · 👤 ' + r.persona : '' }}</small>
                        </span>
                        <button class="btn-outline btn-small" *ngIf="r.en_esta_sede" (click)="abrirDesdeBusqueda(r)">Abrir</button>
                        <button class="btn-primary btn-small" *ngIf="!r.en_esta_sede && capBuscar.elegido!==r" (click)="elegirTraer(r)">Traer aquí</button>
                      </div>
                      <div class="busq-traer" *ngIf="capBuscar.elegido===r">
                        <label>Área en {{ cap?.lev?.sede }} *</label>
                        <select [(ngModel)]="capBuscar.idUbicacion">
                          <option [ngValue]="null">Selecciona el área</option>
                          <option *ngFor="let a of capOpciones.areas" [ngValue]="a.id">{{ a.nombre }}</option>
                        </select>
                        <label>Persona <span class="muted">(de quién es el resguardo)</span></label>
                        <select [(ngModel)]="capBuscar.idPersona">
                          <option [ngValue]="null">Sin cambiar ({{ r.persona || 'sin resguardante' }})</option>
                          <option *ngFor="let x of capOpciones.personas" [ngValue]="x.id">{{ x.nombre }}</option>
                        </select>
                        <div class="busq-btns">
                          <button class="btn-outline" [disabled]="capBuscar.guardando" (click)="capBuscar.elegido=null">Cancelar</button>
                          <button class="btn-primary" [disabled]="capBuscar.guardando" (click)="confirmarTraer(r)">
                            <span class="mini-ring light" *ngIf="capBuscar.guardando"></span>
                            Mover a {{ cap?.lev?.sede }}
                          </button>
                        </div>
                      </div>
                    </div>
                  </div>
                  <div class="busq-vacio" *ngIf="capBuscar.buscado && !capBuscar.cargando && !capBuscar.resultados.length">
                    <p>No está en ninguna sede.</p>
                    <button class="btn-outline" (click)="nuevoDesdeBusqueda()">＋ Registrarlo como bien nuevo</button>
                  </div>
                </div>
                <div class="inline-error" *ngIf="capBuscar.error" style="margin:0 24px 12px">{{ capBuscar.error }}</div>
                <div class="modal-footer">
                  <button class="btn-outline" [disabled]="capBuscar.guardando" (click)="capBuscar=null">Cerrar</button>
                </div>
              </div>
            </div>

            <!-- ===== Asistente: nuevo levantamiento ===== -->
            <div class="modal-backdrop" *ngIf="levNuevo">
              <div class="modal-card wizard">
                <div class="modal-header">
                  <div>
                    <h2>Nuevo levantamiento</h2>
                    <p>Paso {{ levNuevo.paso }} de 4 · {{ ['Sede', 'Lugar y fecha', 'Personas', 'Pregunta'][levNuevo.paso - 1] }}</p>
                  </div>
                  <button class="close-button" [disabled]="levGuardando" (click)="levNuevo=null">×</button>
                </div>
                <div class="wiz-steps"><span *ngFor="let n of [1,2,3,4]" [class.on]="levNuevo.paso>=n"></span></div>

                <div class="form-section" *ngIf="levNuevo.paso===1">
                  <label>Sede</label>
                  <select [(ngModel)]="levNuevo.idSede" (ngModelChange)="cargarCandidatos(); ponerDireccionSede()">
                    <option [ngValue]="null">Selecciona la sede</option>
                    <option *ngFor="let s of activos(catalogs.sedes, levNuevo.idSede, 'id_sede')" [ngValue]="s.id_sede">{{ s.nombre_sede }}</option>
                  </select>
                  <label>Año del inventario</label>
                  <input type="number" [(ngModel)]="levNuevo.anio" min="2020" max="2100">
                </div>

                <div class="form-section" *ngIf="levNuevo.paso===2">
                  <label>Ciudad</label>
                  <input [(ngModel)]="levNuevo.ciudad" placeholder="Ej. Celaya, Guanajuato">
                  <label>Dirección</label>
                  <textarea rows="2" [(ngModel)]="levNuevo.direccion"></textarea>
                  <label>Fecha y hora de inicio</label>
                  <input type="datetime-local" [(ngModel)]="levNuevo.fechaInicio">
                </div>

                <div class="form-section" *ngIf="levNuevo.paso===3">
                  <label>¿Quién va a capturar con el celular?</label>
                  <select [(ngModel)]="levNuevo.idUsuarioCapturista">
                    <option [ngValue]="null">{{ candidatos.length ? 'Selecciona una persona' : 'Cargando…' }}</option>
                    <option *ngFor="let c of candidatos" [ngValue]="c.idUsuario">{{ c.nombre }} · {{ c.rol }}</option>
                  </select>
                  <label>Participantes del acta <span class="muted">(observan y firman)</span></label>
                  <div class="wiz-people">
                    <div *ngFor="let p of levNuevo.participantes; let i = index">
                      <input [(ngModel)]="p.nombre" placeholder="Nombre">
                      <input [(ngModel)]="p.cargo" placeholder="Cargo">
                      <button class="close-button" title="Quitar" (click)="levNuevo.participantes.splice(i, 1)">×</button>
                    </div>
                  </div>
                  <button class="lev-more" *ngIf="levNuevo.participantes.length < 8" (click)="levNuevo.participantes.push({ nombre: '', cargo: '' })">＋ Agregar participante</button>
                </div>

                <div class="form-section" *ngIf="levNuevo.paso===4">
                  <label>Pregunta para cada bien</label>
                  <input [(ngModel)]="levNuevo.pregunta.texto">
                  <label>Respuestas</label>
                  <div class="wiz-ops">
                    <span *ngFor="let o of levNuevo.pregunta.opciones; let i = index">{{ o }}<button (click)="levNuevo.pregunta.opciones.splice(i, 1)" *ngIf="levNuevo.pregunta.opciones.length > 2">×</button></span>
                  </div>
                  <div class="mail-add" *ngIf="levNuevo.pregunta.opciones.length < 6">
                    <input [(ngModel)]="levNuevo.nuevaOpcion" placeholder="Otra respuesta" (keyup.enter)="agregarOpcion()">
                    <button class="btn-outline" (click)="agregarOpcion()">Agregar</button>
                  </div>
                  <div class="wiz-summary">
                    Al crear se genera el <b>Acta de inicio</b> y queda guardada en Documentos.
                  </div>
                </div>

                <div class="inline-error" *ngIf="levError" style="margin:0 24px 12px">{{ levError }}</div>
                <div class="modal-footer">
                  <button class="btn-outline" *ngIf="levNuevo.paso>1" [disabled]="levGuardando" (click)="levNuevo.paso = levNuevo.paso - 1; levError=''">Atrás</button>
                  <button class="btn-primary" *ngIf="levNuevo.paso<4" (click)="siguientePasoLev()">Siguiente</button>
                  <button class="btn-primary" *ngIf="levNuevo.paso===4" [disabled]="levGuardando" (click)="crearLevantamiento()">
                    <span class="mini-ring light" *ngIf="levGuardando"></span>
                    {{ levGuardando ? 'Creando…' : 'Crear y generar acta de inicio' }}
                  </button>
                </div>
              </div>
            </div>
          </section>

          <!-- DOCUMENTOS -->
          <section *ngIf="section==='documentos'">
            <div class="inv-head">
              <div>
                <button class="inv-back" *ngIf="docsVista!=='menu'" (click)="docsVista='menu'; docsBusqueda=''">← Documentos</button>
                <h1>{{ docsVista==='menu' ? 'Documentos' : (docsVista==='ALTA' ? 'Formatos de alta' : docsVista==='BAJA' ? 'Formatos de baja' : docsVista==='RESGUARDO' ? 'Resguardos individuales' : 'Actas de inventario') }}</h1>
                <p class="inv-resumen">
                  {{ isAdmin || canAuthorize ? 'Formatos autorizados de tu sede.' : 'Tus formatos autorizados.' }}
                  Se generan solos al autorizar y se abren aquí mismo.
                </p>
              </div>
              <div class="heading-actions">
                <button class="btn-outline" [disabled]="docsCargando" (click)="loadDocumentos()">⟳ {{ docsCargando ? 'Actualizando…' : 'Actualizar' }}</button>
              </div>
            </div>

            <div class="inventory-panel placeholder-panel" *ngIf="docsError && !documentosLista.length">
              <div class="placeholder-icon">⚠</div>
              <h2>No fue posible cargar los documentos</h2>
              <p>{{ docsError }}</p>
              <button class="btn-primary" (click)="loadDocumentos()">Reintentar</button>
            </div>

            <!-- Tarjetas -->
            <div class="inv-tiles doc-tiles" *ngIf="docsVista==='menu'">
              <button class="inv-tile t-alta" (click)="docsVista='ALTA'">
                <svg class="tile-bg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"/><path d="M14 3v6h6M12 12v6M9 15h6"/></svg>
                <span class="tile-ico"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"/><path d="M14 3v6h6M12 12v6M9 15h6"/></svg></span>
                <span class="tile-text"><strong>Alta</strong><small>Formato FA-04-02-03</small></span>
                <span class="tile-num">{{ docsCargando && !documentosLista.length ? '…' : conteoDocs('ALTA') }}<em>documentos</em></span>
                <span class="tile-go">Ver altas →</span>
              </button>
              <button class="inv-tile t-baja" (click)="docsVista='BAJA'">
                <svg class="tile-bg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"/><path d="M14 3v6h6M9 15h6"/></svg>
                <span class="tile-ico"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"/><path d="M14 3v6h6M9 15h6"/></svg></span>
                <span class="tile-text"><strong>Baja</strong><small>Formato FA-04-02-04</small></span>
                <span class="tile-num">{{ docsCargando && !documentosLista.length ? '…' : conteoDocs('BAJA') }}<em>documentos</em></span>
                <span class="tile-go">Ver bajas →</span>
              </button>
              <button class="inv-tile t-actas" (click)="docsVista='ACTAS'">
                <svg class="tile-bg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M9 4h6l1 2h3v15H5V6h3z"/><path d="M9 12l2 2 4-4"/></svg>
                <span class="tile-ico"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M9 4h6l1 2h3v15H5V6h3z"/><path d="M9 12l2 2 4-4"/></svg></span>
                <span class="tile-text"><strong>Actas de inventario</strong><small>Inicio y cierre del inventario anual</small></span>
                <span class="tile-num">{{ docsCargando && !documentosLista.length ? '…' : conteoActas }}<em>actas</em></span>
                <span class="tile-go">Ver actas →</span>
              </button>
              <div class="inv-tile t-transfer disabled">
                <span class="tile-ico"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 8h13l-3-3M20 16H7l3 3"/></svg></span>
                <span class="tile-text"><strong>Transferencia</strong><small>Cambio de ubicación o responsable</small></span>
                <span class="tile-soon">Próximamente</span>
              </div>
              <button class="inv-tile t-resguardo" (click)="docsVista='RESGUARDO'">
                <svg class="tile-bg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M12 3l8 4v5c0 5-3.5 8-8 9-4.5-1-8-4-8-9V7z"/><path d="M9 12l2 2 4-4"/></svg>
                <span class="tile-ico"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3l8 4v5c0 5-3.5 8-8 9-4.5-1-8-4-8-9V7z"/><path d="M9 12l2 2 4-4"/></svg></span>
                <span class="tile-text"><strong>Resguardo individual</strong><small>Formato FA-04-02-02 · bienes de cada persona</small></span>
                <span class="tile-num">{{ docsCargando && !documentosLista.length ? '…' : conteoDocs('RESGUARDO') }}<em>resguardos</em></span>
                <span class="tile-go">Ver resguardos →</span>
              </button>
            </div>

            <!-- Lista de documentos -->
            <div class="inventory-panel" *ngIf="docsVista!=='menu'">
              <div class="inv-toolbar">
                <div class="search-box">
                  <span>⌕</span>
                  <input [(ngModel)]="docsBusqueda" placeholder="Buscar por folio, clave, bien, sede o persona...">
                  <button class="clear-search" *ngIf="docsBusqueda" (click)="docsBusqueda=''">×</button>
                </div>
                <span class="people-count">{{ documentosFiltrados.length | number }} {{ documentosFiltrados.length === 1 ? 'documento' : 'documentos' }}</span>
              </div>
              <div class="doc-list">
                <button class="doc-row" *ngFor="let d of documentosFiltrados" [class.anulado]="d.estado==='ANULADO'"
                        [disabled]="abriendoDocumento" (click)="abrirDocumentoGuardado(d.idDocumento)">
                  <span class="doc-row-ico" [class.baja]="d.tipo==='BAJA'" [class.acta]="d.tipo && d.tipo.startsWith('ACTA')" [class.resg]="d.tipo==='RESGUARDO'">
                    <svg *ngIf="d.tipo!=='RESGUARDO'" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"/><path d="M14 3v6h6"/></svg>
                    <svg *ngIf="d.tipo==='RESGUARDO'" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3l8 4v5c0 5-3.5 8-8 9-4.5-1-8-4-8-9V7z"/><path d="M9 12l2 2 4-4"/></svg>
                  </span>
                  <span class="doc-row-main">
                    <span class="doc-row-top">
                      <b class="chip-key">{{ d.folio }}</b>
                      <span class="doc-anulado" *ngIf="d.estado==='ANULADO'">ANULADO</span>
                      <span class="doc-estado borrador" *ngIf="d.estado==='BORRADOR'">BORRADOR</span>
                      <span class="doc-estado" *ngIf="d.estado==='HISTORICO'">HISTÓRICO</span>
                    </span>
                    <ng-container *ngIf="d.tipo==='RESGUARDO'">
                      <strong>{{ d.persona }}</strong>
                      <small>{{ d.totalBienes | number }} {{ d.totalBienes === 1 ? 'bien' : 'bienes' }} · {{ d.levantamiento }}</small>
                    </ng-container>
                    <ng-container *ngIf="d.tipo!=='RESGUARDO'">
                      <strong>{{ d.clave }} · {{ d.descripcion }}</strong>
                      <small>{{ d.sede || 'Sin sede' }} · Solicitó {{ d.solicito || '—' }} · Autorizó {{ d.autorizo || '—' }}</small>
                    </ng-container>
                  </span>
                  <span class="doc-row-date">
                    {{ d.fechaCreacion | date:'dd/MM/yyyy' }}
                    <small>{{ d.fechaCreacion | date:'HH:mm' }}</small>
                  </span>
                  <span class="doc-row-go">
                    <span class="mini-ring" *ngIf="abriendoDocumento===d.idDocumento"></span>
                    <ng-container *ngIf="abriendoDocumento!==d.idDocumento">Ver →</ng-container>
                  </span>
                </button>
                <div class="inv-empty" *ngIf="!documentosFiltrados.length && !docsCargando">
                  {{ docsBusqueda ? 'Ningún documento coincide con la búsqueda.' : docsVista==='RESGUARDO' ? 'Aún no hay resguardos. Se arman solos durante el levantamiento de inventario con los bienes revisados de cada persona.' : (docsVista==='ACTAS' ? 'Aún no hay actas. Se generan al iniciar y al cerrar un inventario anual.' : 'Aún no hay formatos de ' + (docsVista==='ALTA' ? 'alta' : 'baja') + '. Se generan solos cuando se autoriza un movimiento.') }}
                </div>
              </div>
            </div>
          </section>

          <!-- CATÁLOGOS -->
          <section *ngIf="section==='catalogos'">
            <div class="page-heading">
              <div>
                <div class="eyebrow">ADMINISTRACIÓN</div>
                <h1>Catálogos</h1>
                <p>Administra la información base que se usa al registrar bienes, personas y ubicaciones.</p>
              </div>
            </div>

            <div class="catalog-layout">
              <!-- LISTA DE CATÁLOGOS -->
              <nav class="catalog-nav">
                <ng-container *ngFor="let g of catalogGroups">
                  <div class="catalog-group">{{ g.label }}</div>
                  <button *ngFor="let c of g.items" [class.selected]="catalog===c.key" (click)="selectCatalog(c.key)">
                    <span>{{ c.label }}</span>
                    <span class="tab-count">{{ (catalogs[c.key] || []).length }}</span>
                  </button>
                </ng-container>
              </nav>

              <!-- REGISTROS DEL CATÁLOGO -->
              <div class="inventory-panel catalog-panel">
                <div class="inventory-toolbar">
                  <div>
                    <h2>{{ catalogConfig.label }}</h2>
                    <p>{{ catalogConfig.descripcion }}</p>
                  </div>
                  <button class="btn-primary" (click)="newCatalogItem()"><span>＋</span> Nuevo registro</button>
                </div>

                <div class="search-row">
                  <div class="search-box">
                    <span>⌕</span>
                    <input [(ngModel)]="catalogFilter" (ngModelChange)="catalogPage=1" [placeholder]="'Buscar en ' + catalogConfig.label.toLowerCase() + '...'">
                  </div>
                  <label class="inline-check" *ngIf="catalogConfig.estatus">
                    <input type="checkbox" [(ngModel)]="catalogShowInactive" (ngModelChange)="catalogPage=1"> Mostrar inactivos
                  </label>
                  <span class="results-count">{{ filteredCatalogRows.length }} registros</span>
                </div>

                <div class="inline-error" *ngIf="catalogError">{{ catalogError }}</div>

                <div class="table-container">
                  <table>
                    <thead>
                      <tr>
                        <th *ngFor="let col of catalogConfig.columnas">{{ col.label | uppercase }}</th>
                        <th *ngIf="catalogConfig.estatus">ESTADO</th>
                        <th class="action-cell">ACCIONES</th>
                      </tr>
                    </thead>
                    <tbody>
                      <tr *ngFor="let x of pagedCatalogRows" [class.inactive-row]="x.estatus === false">
                        <td *ngFor="let col of catalogConfig.columnas; let first = first">
                          <strong *ngIf="first" class="catalog-main">{{ valorColumna(x, col) }}</strong>
                          <span *ngIf="!first" [class.muted]="valorColumna(x, col)==='—'">{{ valorColumna(x, col) }}</span>
                        </td>
                        <td *ngIf="catalogConfig.estatus">
                          <span class="status" [ngClass]="x.estatus === false ? 'status-neutral' : 'status-good'">{{ x.estatus === false ? 'Inactivo' : 'Activo' }}</span>
                        </td>
                        <td class="action-cell">
                          <div class="row-actions">
                            <button title="Editar" (click)="editCatalogItem(x)">✎</button>
                            <button *ngIf="!catalogConfig.estatus || x.estatus !== false"
                                    [title]="catalogConfig.estatus ? 'Desactivar' : 'Eliminar'"
                                    class="delete-action" (click)="deleteCatalogItem(x)">⌫</button>
                            <button *ngIf="catalogConfig.estatus && x.estatus === false" title="Reactivar" (click)="reactivateCatalogItem(x)">↺</button>
                          </div>
                        </td>
                      </tr>
                      <tr *ngIf="!filteredCatalogRows.length">
                        <td [attr.colspan]="catalogConfig.columnas.length + 2" class="empty-state">
                          <div class="empty-icon">⚙</div>
                          <strong>{{ catalogFilter ? 'Ningún registro coincide con la búsqueda' : 'No hay registros en este catálogo' }}</strong>
                          <p *ngIf="!catalogFilter">Agrega el primero con el botón "Nuevo registro".</p>
                        </td>
                      </tr>
                    </tbody>
                  </table>
                </div>

                <div class="table-footer">
                  <span>Mostrando {{ pagedCatalogRows.length }} de {{ filteredCatalogRows.length }} registros</span>
                  <div class="pager" *ngIf="catalogPages > 1">
                    <button class="btn-outline btn-small" [disabled]="catalogPage===1" (click)="catalogPage=catalogPage-1">‹ Anterior</button>
                    <span>Página {{ catalogPage }} de {{ catalogPages }}</span>
                    <button class="btn-outline btn-small" [disabled]="catalogPage===catalogPages" (click)="catalogPage=catalogPage+1">Siguiente ›</button>
                  </div>
                </div>
              </div>
            </div>

            <!-- FORMULARIO DE CATÁLOGO -->
            <div class="modal-backdrop" *ngIf="catalogModal">
              <div class="modal-card">
                <div class="modal-header">
                  <div>
                    <h2>{{ catalogEditId ? 'Editar' : 'Nuevo' }} · {{ catalogConfig.singular }}</h2>
                    <p>Los campos marcados con * son obligatorios.</p>
                  </div>
                  <button class="close-button" [disabled]="catalogSaving" (click)="catalogModal=false">×</button>
                </div>

                <div class="form-section">
                  <div class="form-grid">
                    <div *ngFor="let f of catalogConfig.campos" [class.full-span]="f.tipo==='check' || f.ancho==='completo'">
                      <ng-container [ngSwitch]="f.tipo">
                        <label *ngSwitchCase="'check'" class="inline-check">
                          <input type="checkbox" [(ngModel)]="catForm[f.key]"> {{ f.label }}
                        </label>
                        <ng-container *ngSwitchCase="'select'">
                          <label>{{ f.label }}{{ f.requerido ? ' *' : '' }}</label>
                          <select [(ngModel)]="catForm[f.key]" [class.input-error]="catalogFieldError===f.key">
                            <option [ngValue]="null">Seleccionar</option>
                            <option *ngFor="let o of opcionesCampo(f)" [ngValue]="o.valor">{{ o.texto }}</option>
                          </select>
                        </ng-container>
                        <ng-container *ngSwitchDefault>
                          <label>{{ f.label }}{{ f.requerido ? ' *' : '' }}</label>
                          <input [type]="f.tipo || 'text'" [(ngModel)]="catForm[f.key]" [placeholder]="f.placeholder || ''"
                                 [class.input-error]="catalogFieldError===f.key" [attr.maxlength]="f.max || null"
                                 (keyup.enter)="saveCatalog()">
                        </ng-container>
                      </ng-container>
                    </div>
                  </div>
                  <div class="inline-error" *ngIf="catalogFormError">{{ catalogFormError }}</div>
                </div>

                <div class="modal-footer">
                  <button class="btn-outline" [disabled]="catalogSaving" (click)="catalogModal=false">Cancelar</button>
                  <button class="btn-primary" [disabled]="catalogSaving" (click)="saveCatalog()">
                    {{ catalogSaving ? 'Guardando...' : (catalogEditId ? 'Guardar cambios' : 'Agregar registro') }}
                  </button>
                </div>
              </div>
            </div>
          </section>

        </div>
      </ng-container>
      <div class="foto-visor" *ngIf="fotoAmpliada" (click)="fotoAmpliada=null">
        <img [src]="fotoAmpliada.url" [alt]="fotoAmpliada.titulo" (click)="$event.stopPropagation()">
        <div class="foto-visor-pie"><b>{{ fotoAmpliada.titulo }}</b><button type="button" (click)="fotoAmpliada=null">Cerrar ✕</button></div>
      </div>
    </main>
  </div>
  `,
  styles: [`
    :host{display:block;min-height:100vh;font-family:Inter,Arial,sans-serif;color:#243247}
    *{box-sizing:border-box}
    .app-shell{min-height:100vh;background:#f4f6fa;display:flex}
    .sidebar{position:fixed;top:0;left:0;bottom:0;z-index:95;width:270px;max-width:86vw;overflow-y:auto;background:#17263d;color:#fff;display:flex;flex-direction:column;padding:25px 17px;transform:translateX(0);box-shadow:12px 0 40px #0b172840;transition:transform .28s cubic-bezier(.4,0,.2,1),box-shadow .28s}
    .sidebar.collapsed{transform:translateX(-100%);box-shadow:none}
    .sidebar-backdrop{position:fixed;top:0;right:0;bottom:0;left:0;inset:0;z-index:94;background:#0f1c2e66;backdrop-filter:blur(2px);animation:fadeIn .2s ease both}
    .sidebar-close{margin-left:auto;align-self:flex-start;width:32px;height:32px;border:0;border-radius:8px;background:transparent;color:#a9b8cc;font-size:24px;line-height:1;cursor:pointer;flex-shrink:0}
    .sidebar-close:hover{background:#2b405e;color:#fff}
    .menu-toggle{position:relative;width:40px;height:40px;border-radius:10px;display:grid;place-items:center}
    .toggle-badge{position:absolute;top:2px;right:0;min-width:17px;height:17px;padding:0 4px;border-radius:9px;background:#d9484f;color:#fff;font-size:12px;font-weight:800;line-height:17px;text-align:center;box-shadow:0 0 0 2px #fff}
    .menu-toggle:hover{background:#f1f4f8;color:#17263d}
    .brand{display:flex;align-items:center;gap:12px;margin-bottom:37px}
    .brand-icon{width:39px;height:39px;background:#fff;color:#193452;border-radius:11px;display:grid;place-items:center;font-size:25px;font-weight:900}
    .brand-text{display:flex;flex-direction:column;gap:4px;min-width:0}
    .brand-text strong{font-size:20px;letter-spacing:1.2px}
    .brand-text small{font-size:12px;color:#b9c6d7;white-space:nowrap}
    .menu-label{font-size:12px;letter-spacing:1.4px;color:#8192aa;font-weight:700;margin:5px 10px 12px}
    .nav-item{width:100%;border:0;background:transparent;color:#c1ccdb;display:flex;align-items:center;gap:14px;padding:13px 14px;border-radius:9px;text-align:left;margin-bottom:6px;font-size:15px;cursor:pointer}
    .nav-item:hover,.nav-item.active{background:#2b405e;color:#fff}
    .nav-item.active{box-shadow:inset 3px 0 #5e9be9}
    .nav-icon{font-size:21px;width:22px;text-align:center}
    .sidebar-bottom{margin-top:auto;border-top:1px solid #35455d;padding-top:18px;display:flex;align-items:center;gap:10px}
    .user-avatar,.top-avatar{width:36px;height:36px;border-radius:50%;background:#dce8f7;color:#244568;display:grid;place-items:center;font-size:14px;font-weight:800;flex-shrink:0}
    .user-info{display:flex;flex-direction:column;gap:4px;min-width:0;flex:1}
    .user-info strong{font-size:13px;overflow:hidden;text-overflow:ellipsis}
    .user-info small{font-size:12px;color:#9eacc0}
    .logout{background:transparent;color:#cbd5e1;border:0;font-size:22px;cursor:pointer}
    .main-content{flex:1;min-width:0}
    .topbar{height:76px;background:#fff;border-bottom:1px solid #e8edf3;display:flex;align-items:center;padding:0 32px;gap:18px}
    .menu-toggle{background:transparent;border:0;font-size:22px;color:#607087;cursor:pointer}
    .topbar-title{display:flex;flex-direction:column;gap:5px}
    .topbar-title small{font-size:12px;color:#8793a5;letter-spacing:1px}
    .topbar-title strong{font-size:15.5px;color:#25374e}
    .user-menu-wrap{position:relative}
    .top-avatar-btn{border:0;cursor:pointer;font-family:inherit;transition:box-shadow .2s,transform .15s}
    .top-avatar-btn:hover,.top-avatar-btn.open{box-shadow:0 0 0 3px #244c7833;transform:translateY(-1px)}
    .user-menu{position:absolute;right:0;top:calc(100% + 8px);z-index:45;width:300px;background:#fff;border:1px solid #e3e9f1;border-radius:14px;box-shadow:0 16px 44px #1b2c4430;padding:6px;animation:fadeUp .15s ease both}
    .user-menu-head{display:flex;align-items:center;gap:11px;padding:10px 10px 12px;border-bottom:1px solid #edf1f6;margin-bottom:4px}
    .user-menu-head strong{display:block;font-size:14px;color:#25364c;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:220px}
    .user-menu-head small{display:block;font-size:12px;color:#8d99a9;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:220px}
    .user-menu button{width:100%;display:flex;gap:12px;align-items:center;text-align:left;border:0;background:transparent;padding:10px;border-radius:9px;cursor:pointer;font-family:inherit}
    .user-menu button:hover{background:#f3f6fa}
    .user-menu button strong{display:block;font-size:14px;color:#25364c}
    .user-menu button small{display:block;font-size:12px;color:#8d99a9;margin-top:2px}
    .user-menu-exit strong{color:#bd4b52!important}
    .user-menu-exit .more-ico{background:#fdeced;color:#c2414b}
    .topbar-right{margin-left:auto;display:flex;align-items:center;gap:9px}
    .online-dot{width:8px;height:8px;border-radius:50%;background:#25a878}
    .online-label{font-size:13px;color:#6b7b90}
    .content-area{padding:30px 34px;max-width:1700px;margin:auto}
    .page-heading{display:flex;justify-content:space-between;align-items:center;gap:20px;margin-bottom:25px}
    .eyebrow{font-size:12px;font-weight:800;color:#6482a6;letter-spacing:1.4px;margin-bottom:8px}
    h1{font-size:29px;margin:0 0 7px;font-weight:750;letter-spacing:-.7px;color:#1c2d43}
    .page-heading p{margin:0;color:#8491a3;font-size:15px}
    .heading-actions{display:flex;gap:10px}
    button{font-family:inherit}
    .btn-primary,.btn-outline{border-radius:8px;padding:11px 16px;font-size:14px;font-weight:700;cursor:pointer;display:inline-flex;align-items:center;justify-content:center;gap:8px;transition:.2s}
    .btn-primary{background:#244c78;color:#fff;border:1px solid #244c78}
    .btn-primary:hover{background:#193b63}
    .btn-primary:disabled{opacity:.55;cursor:not-allowed}
    .btn-outline{background:#fff;color:#344c68;border:1px solid #dce3ec}
    .btn-outline:hover{background:#f4f7fb;border-color:#b7c6d9}
    .full{width:100%;margin-top:20px}
    .stats-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:15px;margin-bottom:23px}
    .stat-card{background:#fff;border:1px solid #e8edf3;border-radius:11px;padding:17px 19px;box-shadow:0 2px 7px #253b5507}
    .stat-top{display:flex;align-items:center;justify-content:space-between;color:#718096;font-size:14px}
    .stat-card>strong{display:block;font-size:29px;color:#1c2e45;margin:12px 0 4px;letter-spacing:-.8px}
    .stat-card>small{color:#9aa5b4;font-size:12px}
    .stat-icon{width:31px;height:31px;border-radius:8px;display:grid;place-items:center;font-weight:800}
    .blue{background:#e9f1fc;color:#4677b5}.purple{background:#f0edfc;color:#8068bd}.green{background:#e7f6ef;color:#26936a}.orange{background:#fff3e6;color:#c88a39}
    .inventory-panel{background:#fff;border:1px solid #e7ecf2;border-radius:12px;overflow:hidden;box-shadow:0 3px 12px #233c5708}
    .inventory-toolbar{padding:21px 23px 15px;display:flex;justify-content:space-between;align-items:center;gap:15px}
    .inventory-toolbar h2{font-size:16.5px;margin:0 0 5px;color:#25364c}
    .inventory-toolbar p{font-size:13px;color:#929eae;margin:0}
    .table-actions select{width:auto;min-width:155px}
    .search-row{padding:0 23px 17px;display:flex;align-items:center;gap:15px}
    .search-box{display:flex;align-items:center;gap:10px;border:1px solid #e1e7ef;border-radius:8px;padding:0 12px;height:39px;max-width:490px;flex:1}
    .search-box span{font-size:24px;color:#8493a6}
    .search-box input{border:0!important;outline:0!important;padding:0!important;box-shadow:none!important}
    input,select,textarea{width:100%;border:1px solid #dce3ec;border-radius:7px;padding:10px 11px;font-size:14px;color:#34465d;background:#fff;outline:none;font-family:inherit}
    input:focus,select:focus,textarea:focus{border-color:#7297c3;box-shadow:0 0 0 3px #7297c31c}
    .results-count{font-size:13px;color:#8491a2;white-space:nowrap}
    .table-container{width:100%;overflow:auto}
    table{width:100%;border-collapse:collapse;white-space:nowrap;text-align:left}
    thead{background:#f7f9fc}
    th{font-size:11.5px;letter-spacing:.65px;color:#8491a2;font-weight:800;padding:13px 15px;border-top:1px solid #edf0f4;border-bottom:1px solid #edf0f4}
    td{font-size:13px;color:#526176;padding:13px 15px;border-bottom:1px solid #eff2f6}
    tbody tr:hover{background:#fafbfd}
    .asset-key{color:#365d89;font-size:13px}
    .asset-name{display:flex;flex-direction:column;gap:5px}
    .asset-name strong{font-size:13px;color:#2b3d53}
    .asset-name small{font-size:12px;color:#9ba6b5}
    .quantity{background:#f1f4f8;padding:5px 9px;border-radius:5px;font-weight:700;color:#42566e}
    .status{font-size:11.5px;font-weight:800;padding:6px 9px;border-radius:20px;display:inline-block}
    .status-good{background:#e8f6ee;color:#21835a}
    .status-warning{background:#fff3df;color:#b27b22}
    .status-bad{background:#fdeced;color:#bd5056}
    .status-neutral{background:#edf1f6;color:#63738a}
    .row-actions{display:flex;gap:5px}
    .row-actions button{width:27px;height:27px;border:1px solid #e5eaf0;background:#fff;border-radius:6px;color:#60738a;cursor:pointer}
    .row-actions button:hover{background:#eef4fa}
    .row-actions .delete-action{color:#c65d62}
    .table-footer{display:flex;justify-content:space-between;padding:15px 22px;color:#8995a5;font-size:12px}
    .empty-state{text-align:center!important;padding:42px!important;color:#8491a2!important}
    .empty-state p{font-size:13px}
    .empty-icon{font-size:30px;color:#a8b6c7;margin-bottom:10px}
    .modal-backdrop{position:fixed;top:0;right:0;bottom:0;left:0;inset:0;background:#15243b88;z-index:100;display:flex;align-items:center;justify-content:center;padding:20px}
    .modal-card{width:100%;max-width:590px;max-height:90vh;overflow:auto;background:#fff;border-radius:14px;box-shadow:0 20px 70px #15243b35}
    .large-modal{max-width:850px}
    .excel-modal{max-width:760px}
    .modal-header{padding:22px 24px;border-bottom:1px solid #edf0f4;display:flex;align-items:flex-start;justify-content:space-between}
    .modal-header h2{margin:0 0 6px;font-size:20px;color:#25364c}
    .modal-header p{margin:0;color:#8b97a7;font-size:13px}
    .close-button{background:transparent;border:0;color:#8290a2;font-size:27px;cursor:pointer}
    .form-section{padding:22px 24px}
    .form-section h3,.catalog-form h2{font-size:15px;color:#30445d;margin:0 0 17px}
    .form-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px 17px;margin-bottom:17px}
    .form-grid label,.form-section label,.login-card label{display:block;font-size:12px;font-weight:700;color:#64748a;margin-bottom:7px}
    .form-section textarea{resize:vertical}
    .modal-footer{display:flex;justify-content:flex-end;gap:10px;padding:16px 24px;border-top:1px solid #edf0f4}
    .excel-instructions{display:flex;gap:13px;padding:20px 24px 12px}
    .excel-icon{width:39px;height:39px;background:#e7f5ed;color:#21865d;border-radius:9px;display:grid;place-items:center;font-size:24px}
    .excel-instructions strong{font-size:14px;color:#33465e}
    .excel-instructions p,.preview-heading p{font-size:12px;color:#8995a5;margin:5px 0 0}
    .upload-zone{margin:10px 24px 20px;border:1.5px dashed #cbd7e4;border-radius:11px;min-height:165px;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:8px;background:#fafcff;cursor:pointer}
    .upload-zone:hover{border-color:#638dbb;background:#f5f9ff}
    .upload-zone.has-file{border-color:#58a77d;background:#f6fcf8}
    .upload-zone input{display:none}
    .upload-symbol{font-size:28px;color:#6486ac}
    .upload-zone strong{font-size:14px;color:#354b65}
    .upload-zone small{font-size:12px;color:#91a0b1}
    .upload-button{padding:8px 13px;margin-top:4px;font-size:12px}
    .excel-error{margin:0 24px 15px;color:#bb454b;font-size:13px}
    .preview-section{padding:0 24px 20px}
    .preview-heading{display:flex;justify-content:space-between;align-items:center;margin-bottom:12px}
    .preview-heading strong{font-size:14px;color:#33465e}
    .preview-table{border:1px solid #e8edf3;border-radius:7px;max-height:190px}
    .preview-table th,.preview-table td{padding:9px 11px}
    .preview-note{display:block;color:#8b97a7;font-size:12px;margin-top:10px}
    .detail-list{padding:22px 24px;display:grid;grid-template-columns:1fr 1fr;gap:17px}
    .detail-list div{display:flex;flex-direction:column;gap:5px}
    .detail-list small{font-size:12px;color:#8996a7}
    .detail-list strong{font-size:14px;color:#34485f}
    .placeholder-panel{text-align:center;padding:75px 20px}
    .placeholder-icon{font-size:40px;color:#7c9cbd}
    .placeholder-panel h2{font-size:19px}
    .placeholder-panel p{font-size:14px;color:#8995a5}
    .coming-soon{display:inline-block;margin-top:12px;padding:7px 12px;border-radius:20px;background:#eef3f9;color:#54779d;font-size:12px}
    .catalog-tabs{display:flex;gap:5px;flex-wrap:wrap;padding:18px 22px;border-bottom:1px solid #edf0f4}
    .catalog-tabs button{border:0;background:#f3f6fa;padding:9px 13px;border-radius:7px;font-size:13px;color:#687b91;cursor:pointer}
    .catalog-tabs button.selected{background:#244c78;color:white}
    .catalog-form{padding:22px}
    .catalog-form .form-grid>div{max-width:460px}
    .catalog-form label{display:block;font-size:12px;font-weight:700;color:#64748a;margin:12px 0 6px}
    .login-screen{min-height:100vh;display:grid;place-items:center;background:linear-gradient(135deg,#17263d,#294b70)}
    .login-card{background:#fff;width:min(410px,calc(100% - 35px));padding:38px;border-radius:15px;box-shadow:0 20px 70px #0c1b2d55}
    .login-brand{margin-bottom:30px;color:#243c58}
    .login-card h1{font-size:27px}
    .login-card>p{font-size:14px;color:#8995a5;margin-bottom:25px}
    .login-card label{margin-top:16px}
    .login-card input{height:42px}
    .error-message{color:#c24b52;font-size:13px;margin-top:12px}
    /* ===== INVENTARIO: CARGA ===== */
    /* ---------- INVENTARIO (vista limpia) ---------- */
    .inv-head{display:flex;justify-content:space-between;align-items:flex-end;gap:20px;margin-bottom:20px}
    .inv-head h1{font-size:28px;letter-spacing:-.6px;color:#14243a;margin:0 0 8px}
    .inv-resumen{display:flex;flex-wrap:wrap;align-items:center;gap:6px 18px;margin:0;font-size:14.5px;color:#7d8a9b}
    .inv-resumen b{color:#25364c;font-weight:700}
    .inv-resumen span{position:relative}
    .inv-resumen span+span::before{content:'';position:absolute;left:-10px;top:50%;width:3px;height:3px;border-radius:50%;background:#c3ccd8}
    .inv-scope{font-size:12.5px;font-weight:700;color:#4f6f95;background:#eef3f9;padding:3px 9px;border-radius:20px}
    .inv-resumen span.inv-scope::before{display:none}
    .more-wrap{position:relative}
    .caret{font-size:12px;color:#8a97a8}
    .more-menu{position:absolute;right:0;top:calc(100% + 6px);z-index:40;width:280px;background:#fff;border:1px solid #e3e9f1;border-radius:12px;box-shadow:0 14px 40px #1b2c4426;padding:6px;animation:fadeUp .15s ease both}
    .more-menu button{width:100%;display:flex;gap:12px;align-items:center;text-align:left;border:0;background:transparent;padding:10px;border-radius:9px;cursor:pointer;font-family:inherit}
    .more-menu button:hover:not(:disabled){background:#f3f6fa}
    .more-menu button:disabled{opacity:.45;cursor:not-allowed}
    .more-ico{width:34px;height:34px;border-radius:9px;background:#eef3f9;color:#244c78;display:grid;place-items:center;font-size:16.5px;flex-shrink:0}
    .more-menu strong{display:block;font-size:14.5px;color:#25364c}
    .more-menu small{display:block;font-size:13px;color:#8d99a9;margin-top:2px}
    .class-tabs{display:flex;gap:4px;padding:10px 16px 0;border-bottom:1px solid #edf1f6;overflow-x:auto;scrollbar-width:thin}
    .class-tabs button{flex-shrink:0;display:inline-flex;align-items:center;gap:7px;border:0;background:transparent;font-family:inherit;font-size:14px;font-weight:600;color:#6f7d8f;padding:10px 12px 12px;border-bottom:2px solid transparent;margin-bottom:-1px;cursor:pointer;white-space:nowrap;transition:color .15s,border-color .15s}
    .class-tabs button:hover{color:#25364c}
    .class-tabs button.active{color:#1c3a5e;border-bottom-color:#244c78}
    .class-tabs i{width:7px;height:7px;border-radius:50%;display:inline-block}
    .class-tabs span{font-size:12.5px;font-weight:700;color:#8a97a8;background:#f1f4f8;padding:2px 7px;border-radius:10px}
    .class-tabs button.active span{background:#e5eef8;color:#244c78}
    .inv-toolbar{display:flex;align-items:center;gap:12px;padding:14px 16px}
    .inv-toolbar .search-box{flex:1;min-width:0}
    .inv-toolbar+.search-row{padding:0 16px 12px}
    @media(max-width:700px){.inv-head{flex-direction:column;align-items:stretch}.inv-head .heading-actions>*{flex:1}.inv-toolbar{flex-wrap:wrap}.more-menu{left:0;right:auto}}

    /* ---------- FORMATOS PDF ---------- */
    .aviso-btn-sec{width:100%;height:42px;margin-top:8px;border:0;border-radius:12px;background:transparent;color:#6b7686;font-family:inherit;font-size:15px;font-weight:600;cursor:pointer}
    .aviso-btn-sec:hover{background:#f3f5f8}
    .doc-backdrop{z-index:215}
    .doc-card{max-width:980px!important;width:100%}
    .doc-loading{display:flex;flex-direction:column;align-items:center;gap:14px;padding:60px 20px;color:#41546b}
    .doc-title{display:flex;align-items:center;gap:14px}
    .doc-badge{font-size:13px;font-weight:800;letter-spacing:1.5px;color:#fff;background:#21835a;border-radius:8px;padding:8px 10px}
    .doc-badge.baja{background:#c2414b}
    .doc-body{display:grid;grid-template-columns:minmax(0,1fr) 300px;gap:0;min-height:520px}
    .doc-preview{background:#e9edf2;padding:14px}
    .doc-preview iframe{width:100%;height:100%;min-height:500px;border:0;border-radius:8px;background:#fff;box-shadow:0 4px 18px #1b2c4422}
    .doc-side{padding:18px 20px;border-left:1px solid #edf0f4;display:flex;flex-direction:column;gap:8px}
    .doc-side h3{margin:6px 0 4px;font-size:12.5px;font-weight:800;letter-spacing:1.2px;color:#8a97a8;text-transform:uppercase}
    .doc-side label{font-size:13px;font-weight:700;color:#4a5a6e;margin-top:4px}
    .doc-action{display:flex;align-items:center;gap:12px;text-align:left;border:1px solid #e3e9f1;background:#fff;border-radius:11px;padding:10px 12px;cursor:pointer;font-family:inherit;transition:background .15s,border-color .15s}
    .doc-action:hover{background:#f4f7fb;border-color:#c9d6e6}
    .doc-action>span{width:34px;height:34px;border-radius:9px;background:#eef3f9;color:#244c78;display:grid;place-items:center;font-size:17.5px;flex-shrink:0}
    .doc-action strong{display:block;font-size:14.5px;color:#25364c}
    .doc-action small{display:block;font-size:12.5px;color:#8d99a9;margin-top:2px;word-break:break-all}
    .doc-action em{margin-left:auto;font-style:normal;font-size:22px;color:#a3b0c0}
    .ico-wa{background:#e3f8eb!important}
    .ico-wa svg{width:19px;height:19px;fill:#25d366}
    .doc-file-mini{margin-top:auto;font-size:12.5px;color:#8d99a9;word-break:break-all;padding-top:10px}
    .doc-flow{max-width:620px;margin:0 auto;padding:18px 24px 26px;display:flex;flex-direction:column;gap:16px;animation:fadeUp .25s ease both}
    .doc-back{align-self:flex-start;border:0;background:transparent;padding:4px 0;font-family:inherit;font-size:14.5px;font-weight:700;color:#4f6f95;cursor:pointer}
    .doc-back:hover{color:#17263d;text-decoration:underline}
    .flow-head{display:flex;align-items:center;gap:14px}
    .flow-head h2{margin:0 0 3px;font-size:21px;color:#14243a}
    .flow-head p{margin:0;font-size:14px;color:#7d8a9b;word-break:break-all}
    .flow-ico{width:48px;height:48px;border-radius:14px;background:#eef3f9;color:#244c78;display:grid;place-items:center;flex-shrink:0}
    .flow-ico svg{width:24px;height:24px}
    .flow-step{background:#fafbfd;border:1px solid #edf1f6;border-radius:14px;padding:16px;display:flex;flex-direction:column;gap:10px}
    .step-title{display:flex;align-items:center;gap:9px;font-size:15px;font-weight:800;color:#25364c}
    .step-title span{width:22px;height:22px;border-radius:50%;background:#17263d;color:#fff;display:grid;place-items:center;font-size:13px}
    .step-title em{font-style:normal;font-weight:500;color:#8d99a9;font-size:14px}
    .flow-help{font-size:13px;color:#8d99a9}
    .wa-input{height:46px;font-size:17.5px!important;letter-spacing:1px}
    .wa-steps{margin:0;padding-left:20px;font-size:14.5px;line-height:1.9;color:#41546b}
    .flow-actions{display:flex;justify-content:flex-end;gap:10px;flex-wrap:wrap}
    .flow-actions .wa-btn{padding:0 22px}
    .wa-box{display:flex;flex-direction:column;gap:6px;background:#f2fbf5;border:1px solid #cdeedb;border-radius:12px;padding:12px}
    .wa-box label{margin-top:0!important}
    .wa-btn{display:inline-flex;align-items:center;justify-content:center;gap:9px;height:42px;border:0;border-radius:10px;background:#25d366;color:#fff;font-family:inherit;font-size:15.5px;font-weight:800;cursor:pointer;transition:filter .2s,transform .15s}
    .wa-btn:hover{filter:brightness(.95);transform:translateY(-1px)}
    .wa-btn svg{width:19px;height:19px;fill:#fff}
    .wa-direct{border:0;background:transparent;color:#1d7a46;font-family:inherit;font-size:13.5px;font-weight:700;text-decoration:underline;cursor:pointer;padding:2px}
    .wa-manual{display:block;margin-top:6px;font-weight:800;color:#128c4a;text-decoration:underline}
    .wa-hint{font-size:13.5px;line-height:1.45;color:#1d7a46;background:#fff;border-radius:8px;padding:8px 10px}
    @media(max-width:860px){.flow-actions>*{flex:1}.doc-flow{padding:14px 16px 20px}.doc-body{grid-template-columns:1fr}.doc-preview iframe{min-height:420px}.doc-side{border-left:0;border-top:1px solid #edf0f4}}

    /* ---------- REVOCACIÓN ---------- */
    .revoke-link{border:1px dashed #e6b866;background:#fffaf0;color:#a56a00;font-family:inherit;font-size:14.5px;font-weight:700;border-radius:10px;padding:11px 14px;width:100%;text-align:left;cursor:pointer;transition:background .2s}
    .revoke-link:hover{background:#fff1d6}
    .revoke-box{background:#fffaf0;border:1px solid #f0d49a;border-radius:12px;padding:16px;color:#7a5200;animation:fadeUp .2s ease both}
    .revoke-box>strong{display:block;font-size:15.5px;margin-bottom:8px}
    .revoke-box ul{margin:0 0 14px;padding-left:18px;font-size:14px;line-height:1.7}
    .revoke-box label{color:#7a5200}
    .revoke-box textarea{background:#fff}
    .revoke-off{margin:0;font-size:13.5px;color:#8a97a8}
    .btn-revoke{border-radius:8px;padding:11px 16px;font-size:14px;font-weight:700;cursor:pointer;display:inline-flex;align-items:center;gap:8px;background:#d98a00;color:#fff;border:1px solid #c47c00;transition:.2s}
    .btn-revoke:hover:not(:disabled){background:#bf7900}
    .btn-revoke:disabled{opacity:.6;cursor:not-allowed}
    .revoked-note{background:#fffaf0;border:1px solid #f0d49a;border-radius:10px;padding:14px 16px;color:#7a5200;font-size:14px}
    .revoked-note p{margin:4px 0 8px;font-size:13px;opacity:.85}
    .revoked-tag{display:inline-block;margin-left:6px;font-size:12px;font-weight:800;color:#a56a00;background:#fff1d6;border-radius:20px;padding:3px 8px;vertical-align:middle}

    /* ---------- AVISO DE RESULTADO ---------- */
    .aviso-backdrop{position:fixed;top:0;right:0;bottom:0;left:0;inset:0;z-index:220;background:#0f1c2e55;backdrop-filter:blur(3px);display:grid;place-items:center;padding:16px;animation:fadeIn .2s ease both}
    .aviso-card{--av:#21835a;--av-bg:#e8f6ee;position:relative;overflow:hidden;width:min(380px,100%);background:#fff;border-radius:22px;padding:34px 28px 26px;text-align:center;box-shadow:0 30px 80px #0b172855;animation:avisoEntra .45s cubic-bezier(.2,1.3,.4,1) both}
    .av-pendiente{--av:#c77700;--av-bg:#fff4dc}
    .av-baja{--av:#c2414b;--av-bg:#fdeced}
    .av-rechazo{--av:#c2414b;--av-bg:#fdeced}
    .av-error{--av:#c2414b;--av-bg:#fdeced}
    .av-info{--av:#2d5d93;--av-bg:#eaf1fa}
    .aviso-ico{width:84px;height:84px;margin:0 auto 18px;border-radius:50%;background:var(--av-bg);display:grid;place-items:center;animation:avisoPop .5s .1s cubic-bezier(.2,1.5,.4,1) both}
    .aviso-ico svg{width:58px;height:58px;fill:none;stroke:var(--av);stroke-width:3.2;stroke-linecap:round;stroke-linejoin:round}
    .av-circ{stroke-dasharray:145;stroke-dashoffset:145;animation:avisoTrazo .6s .15s ease forwards;opacity:.35}
    .av-trazo{stroke-dasharray:80;stroke-dashoffset:80;animation:avisoTrazo .45s .5s ease forwards}
    .aviso-card h2{margin:0 0 8px;font-size:24px;font-weight:800;letter-spacing:-.3px;color:#14243a}
    .aviso-card p{margin:0 0 24px;font-size:15.5px;line-height:1.55;color:#5f6f86}
    .aviso-btn{width:100%;height:48px;border:0;border-radius:13px;background:var(--av);color:#fff;font-family:inherit;font-size:15.5px;font-weight:700;cursor:pointer;transition:filter .2s,transform .15s}
    .aviso-btn:hover{filter:brightness(1.08);transform:translateY(-1px)}
    .aviso-timer{position:absolute;left:0;right:0;bottom:0;height:4px;background:var(--av-bg)}
    .aviso-timer span{display:block;height:100%;background:var(--av);transform-origin:left;animation:avisoTiempo linear forwards}
    @keyframes avisoEntra{from{opacity:0;transform:scale(.85) translateY(14px)}to{opacity:1;transform:none}}
    @keyframes avisoPop{from{transform:scale(.4);opacity:0}to{transform:none;opacity:1}}
    @keyframes avisoTrazo{to{stroke-dashoffset:0}}
    @keyframes avisoTiempo{from{transform:scaleX(1)}to{transform:scaleX(0)}}
    @media (prefers-reduced-motion:reduce){.aviso-card,.aviso-ico{animation:none}.av-circ,.av-trazo{animation:none;stroke-dashoffset:0}}

    /* ---------- AUTORIZACIONES: PENDIENTES QUE BRILLAN ---------- */
    @keyframes brilloFila{0%,100%{background:#fffaf0}50%{background:#ffefc9}}
    @keyframes brilloBarra{0%,100%{box-shadow:inset 4px 0 #f0a020}50%{box-shadow:inset 4px 0 #ffcf5c,inset 14px 0 18px -10px #f5b33c88}}
    @keyframes halo{0%,100%{box-shadow:0 0 0 0 #f0a02066,0 3px 12px #f0a0201f}50%{box-shadow:0 0 0 6px #f0a02000,0 6px 22px #f0a02047}}
    @keyframes latido{0%,100%{transform:scale(1)}50%{transform:scale(1.15)}}
    @keyframes destello{0%{left:-60%}60%,100%{left:130%}}
    .row-pendiente td{animation:brilloFila 2.4s ease-in-out infinite}
    .row-pendiente td:first-child{animation:brilloFila 2.4s ease-in-out infinite,brilloBarra 2.4s ease-in-out infinite}
    .row-pendiente:hover td{background:#ffe9b8!important}
    .row-pendiente .status{position:relative;padding-left:20px}
    .row-pendiente .status::before{content:'';position:absolute;left:8px;top:50%;width:7px;height:7px;margin-top:-3.5px;border-radius:50%;background:#e8920c;box-shadow:0 0 0 0 #e8920c88;animation:halo 1.6s ease-out infinite}
    .btn-revisar{position:relative;overflow:hidden;background:#f0a020!important;border-color:#e0920f!important;color:#fff!important;box-shadow:0 4px 12px #f0a02055}
    .btn-revisar:hover{background:#dc8d0c!important}
    .btn-revisar::after{content:'';position:absolute;top:0;bottom:0;left:-60%;width:40%;background:linear-gradient(100deg,transparent,#ffffff8c,transparent);transform:skewX(-20deg);animation:destello 2.6s ease-in-out infinite}
    .stat-alerta{border-color:#f5c26b!important;background:linear-gradient(180deg,#fffaf0,#fff)!important;animation:halo 2.2s ease-in-out infinite}
    .stat-alerta>strong{color:#c77700!important}
    .tab-alerta{background:#f0a020!important;color:#fff!important;animation:latido 1.4s ease-in-out infinite}
    .nav-badge,.toggle-badge{animation:latido 1.4s ease-in-out infinite}
    @media (prefers-reduced-motion:reduce){.row-pendiente td,.row-pendiente td:first-child,.btn-revisar::after,.stat-alerta,.tab-alerta,.nav-badge,.toggle-badge,.row-pendiente .status::before{animation:none}.row-pendiente td{background:#fff4dc}.row-pendiente td:first-child{box-shadow:inset 4px 0 #f0a020}}

    /* ---------- INVENTARIO: TARJETAS ---------- */
    .inv-back{border:0;background:transparent;padding:0;margin:0 0 8px;font-family:inherit;font-size:14.5px;font-weight:700;color:#4f6f95;cursor:pointer}
    .inv-back:hover{color:#17263d;text-decoration:underline}
    .inv-tiles{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:18px;min-height:calc(100vh - 230px)}
    .inv-tile{position:relative;overflow:hidden;border:0;border-radius:20px;padding:28px 30px;text-align:left;font-family:inherit;color:#fff;cursor:pointer;display:flex;flex-direction:column;justify-content:space-between;gap:18px;min-height:220px;box-shadow:0 10px 30px #17263d26;transition:transform .2s ease,box-shadow .2s ease;animation:fadeUp .45s ease both}
    .inv-tile:nth-child(2){animation-delay:.05s}.inv-tile:nth-child(3){animation-delay:.1s}.inv-tile:nth-child(4){animation-delay:.15s}
    .inv-tile:hover{transform:translateY(-4px);box-shadow:0 18px 40px #17263d3d}
    .inv-tile:active{transform:translateY(-1px)}
    .t-sedes{background:linear-gradient(140deg,#17304f,#2d5d93)}
    .t-bienes{background:linear-gradient(140deg,#0f4c47,#1f8a7c)}
    .t-personas{background:linear-gradient(140deg,#3b2a6b,#6a4fb3)}
    .t-completo{background:linear-gradient(140deg,#1f2937,#4b5563)}
    .tile-bg{position:absolute;right:-30px;bottom:-40px;width:230px;height:230px;opacity:.1;pointer-events:none;transition:transform .35s ease}
    .inv-tile:hover .tile-bg{transform:scale(1.08) rotate(-4deg)}
    .tile-ico{width:64px;height:64px;border-radius:18px;background:#ffffff1f;border:1px solid #ffffff2e;display:grid;place-items:center}
    .tile-ico svg{width:34px;height:34px}
    .tile-text strong{display:block;font-size:28px;font-weight:800;letter-spacing:-.4px}
    .tile-text small{display:block;font-size:15px;color:#ffffffb8;margin-top:5px}
    .tile-num{font-size:42px;font-weight:800;letter-spacing:-1px;line-height:1}
    .tile-num em{font-style:normal;font-size:15px;font-weight:600;letter-spacing:0;color:#ffffffb0;margin-left:8px}
    .tile-go{position:absolute;right:26px;top:30px;font-size:14px;font-weight:700;background:#ffffff1f;border:1px solid #ffffff2e;padding:7px 13px;border-radius:20px;transition:background .2s}
    .inv-tile:hover .tile-go{background:#ffffff33}

    .sede-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:16px}
    .sede-card{position:relative;text-align:left;font-family:inherit;background:#fff;border:1px solid #e5eaf1;border-radius:16px;padding:20px;cursor:pointer;display:flex;flex-direction:column;gap:12px;box-shadow:0 3px 12px #233c5708;transition:transform .18s,box-shadow .18s,border-color .18s;animation:fadeUp .35s ease both}
    .sede-card:hover{transform:translateY(-3px);border-color:#bcd0e8;box-shadow:0 12px 28px #233c5717}
    .sede-top{display:flex;gap:12px;align-items:center}
    .sede-ico{width:46px;height:46px;border-radius:13px;background:linear-gradient(140deg,#17304f,#2d5d93);color:#fff;display:grid;place-items:center;flex-shrink:0}
    .sede-ico svg{width:24px;height:24px}
    .sede-top strong{display:block;font-size:16.5px;color:#14243a}
    .sede-top small{display:block;font-size:13.5px;color:#8a97a8;margin-top:3px}
    .sede-num{font-size:36px;font-weight:800;color:#14243a;letter-spacing:-1px;line-height:1}
    .sede-num span{font-size:14px;font-weight:600;color:#8a97a8;letter-spacing:0;margin-left:7px}
    .sede-bar{height:6px;background:#edf1f6;border-radius:4px;overflow:hidden}
    .sede-bar span{display:block;height:100%;background:#2fa36b;border-radius:4px}
    .sede-meta{display:flex;justify-content:space-between;font-size:13.5px;color:#6f7d8f}
    .sede-tags{display:flex;flex-wrap:wrap;gap:6px}
    .sede-tags span{display:inline-flex;align-items:center;gap:6px;font-size:12.5px;font-weight:700;color:#4d5f75;background:#f3f6fa;border-radius:20px;padding:4px 10px}
    .sede-tags i{width:7px;height:7px;border-radius:50%;display:inline-block}
    .sede-go{font-size:14px;font-weight:700;color:#2d5d93;margin-top:2px}

    .people-bar{display:flex;align-items:center;gap:14px;margin-bottom:16px}
    .people-bar .search-box{background:#fff;max-width:420px}
    .people-count{font-size:14px;color:#7d8a9b;font-weight:600}
    .people-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:12px}
    .person-card{display:flex;align-items:center;gap:14px;text-align:left;font-family:inherit;background:#fff;border:1px solid #e5eaf1;border-radius:14px;padding:14px 16px;cursor:pointer;transition:transform .15s,box-shadow .15s,border-color .15s}
    .person-card:hover{transform:translateY(-2px);border-color:#cbbdf0;box-shadow:0 10px 24px #3b2a6b14}
    .person-avatar{width:46px;height:46px;border-radius:50%;background:linear-gradient(140deg,#3b2a6b,#6a4fb3);color:#fff;display:grid;place-items:center;font-size:15.5px;font-weight:800;flex-shrink:0}
    .person-info{flex:1;min-width:0}
    .person-info strong{display:block;font-size:15px;color:#14243a;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
    .person-info small{display:block;font-size:13px;color:#8a97a8;margin-top:2px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
    .person-tags{display:flex;flex-wrap:wrap;gap:5px;margin-top:7px}
    .person-tags span{font-size:12px;font-weight:700;color:#5b47a0;background:#f1edfb;border-radius:20px;padding:3px 8px}
    .person-num{text-align:right;font-size:26px;font-weight:800;color:#3b2a6b;line-height:1;flex-shrink:0}
    .person-num em{display:block;font-style:normal;font-size:12.5px;font-weight:600;color:#8a97a8;margin-top:3px}

    .asset-cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(290px,1fr));gap:14px;padding:4px 16px 16px}
    .asset-card{display:flex;flex-direction:column;gap:7px;text-align:left;font-family:inherit;background:#fff;border:1px solid #e5eaf1;border-left:5px solid #c3ccd8;border-radius:14px;padding:15px 16px;cursor:pointer;transition:transform .15s,box-shadow .15s}
    .asset-card:hover{transform:translateY(-2px);box-shadow:0 10px 24px #233c5714}
    .asset-chips{display:flex;flex-wrap:wrap;gap:6px;align-items:center;margin-bottom:3px}
    .chip-key{font-family:"SFMono-Regular",Consolas,monospace;font-size:13px;font-weight:800;color:#14243a;background:#f3f5f8;border:1px solid #e1e6ed;border-radius:7px;padding:3px 8px}
    .chip-class{font-size:12px;font-weight:800;letter-spacing:.3px;color:#4f6f95;background:#eef3f9;border-radius:7px;padding:4px 8px;text-transform:uppercase;max-width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .asset-title{font-size:16px;color:#14243a;line-height:1.3;text-transform:uppercase}
    .asset-sub{font-size:14px;color:#7d8a9b;margin-top:-3px}
    .asset-line{display:flex;align-items:center;gap:7px;font-size:14px;color:#4b5a6d;min-width:0}
    .asset-line svg{width:14px;height:14px;flex-shrink:0;color:#94a2b4}
    .asset-line.muted{color:#a3adba}
    .asset-foot{display:flex;justify-content:space-between;align-items:center;border-top:1px dashed #e5eaf1;padding-top:9px;margin-top:4px;font-size:13.5px;color:#8a97a8}
    .asset-more{font-weight:700;color:#2d5d93}
    .more-cards{display:flex;align-items:center;justify-content:center;gap:14px;padding:6px 16px 20px;font-size:14px;color:#7d8a9b}
    .inv-empty{grid-column:1/-1;text-align:center;padding:40px 16px;color:#7d8a9b;font-size:15px}
    @media(max-width:900px){.inv-tiles{grid-template-columns:1fr;min-height:0}.inv-tile{min-height:170px;padding:22px}.tile-num{font-size:34px}.tile-text strong{font-size:24px}}
    @media(max-width:700px){.asset-cards{grid-template-columns:1fr;padding:4px 12px 12px}.people-grid,.sede-grid{grid-template-columns:1fr}.people-bar{flex-wrap:wrap}.tile-go{display:none}}

    /* ---------- INVENTARIO ANUAL ---------- */
    .lev-list{display:flex;flex-direction:column;gap:12px}
    .lev-card{background:#fff;border:1px solid #e5eaf1;border-radius:16px;overflow:hidden;box-shadow:0 3px 12px #233c5708;transition:box-shadow .2s,border-color .2s}
    .lev-card.open{border-color:#c9d6e6;box-shadow:0 10px 28px #233c5714}
    .lev-head{width:100%;display:flex;align-items:center;gap:14px;border:0;background:transparent;padding:16px 18px;text-align:left;font-family:inherit;cursor:pointer}
    .lev-ico{width:46px;height:46px;border-radius:13px;background:linear-gradient(140deg,#17304f,#2d5d93);color:#fff;display:grid;place-items:center;font-size:20px;flex-shrink:0}
    .lev-card.cerrado .lev-ico{background:linear-gradient(140deg,#0f4c3a,#21835a)}
    .lev-main{flex:1;min-width:0;display:flex;flex-direction:column;gap:6px}
    .lev-main strong{font-size:16.5px;color:#14243a}
    .lev-main small{font-size:13px;color:#7d8a9b}
    .lev-bar{display:block;height:7px;background:#edf1f6;border-radius:5px;overflow:hidden}
    .lev-bar span{display:block;height:100%;background:linear-gradient(90deg,#2d5d93,#21835a);border-radius:5px;transition:width .4s}
    .lev-chip{font-size:12px;font-weight:800;color:#a56a00;background:#fff4dc;border-radius:20px;padding:5px 11px;flex-shrink:0}
    .lev-chip.ok{color:#21835a;background:#e8f6ee}
    .lev-chev{font-size:20px;color:#9aa6b6;transition:transform .2s;flex-shrink:0}
    .lev-card.open .lev-chev{transform:rotate(180deg)}
    .lev-body{padding:0 18px 18px;display:flex;flex-direction:column;gap:14px;animation:fadeUp .2s ease both}
    .lev-stats{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px}
    .lev-stats div{background:#f6f8fb;border-radius:12px;padding:12px}
    .lev-stats b{display:block;font-size:22px;color:#14243a}
    .lev-stats small{font-size:12.5px;color:#7d8a9b}
    .lev-actions{display:flex;flex-wrap:wrap;gap:10px}
    .cap-avance-btn{align-self:flex-start;border:1px solid #d6e0ec;background:#fff;color:#2d5d93;font-family:inherit;font-size:14px;font-weight:700;border-radius:999px;padding:7px 14px;cursor:pointer}
    .av-aviso{display:flex;flex-direction:column;gap:8px;background:#fff7ec;border:1px solid #f3d6a8;border-radius:14px;padding:14px 16px;color:#7a4b0c;font-size:14px}
    .av-aviso b{font-size:15px}
    .av-aviso .btn-danger{align-self:flex-start}
    .av-chips{display:grid;grid-template-columns:repeat(7,minmax(0,1fr));gap:8px}
    .av-chips .f-noenc b{color:#bd4b52}
    .seg{display:inline-flex;align-self:flex-start;background:#e9eef5;border-radius:999px;padding:4px;gap:4px}
    .seg button{border:0;background:transparent;font-family:inherit;font-size:14px;font-weight:700;color:#55667c;border-radius:999px;padding:8px 14px;cursor:pointer}
    .seg button.on{background:#fff;color:#17304f;box-shadow:0 2px 8px rgba(20,36,58,.12)}
    .st-no .av-ico{background:#fdecee;color:#bd4b52}
    .st-no .av-txt small{color:#bd4b52}
    .av-orig{font-style:normal;font-size:11.5px;font-weight:800;color:#6a4fb3;background:#f1edfb;border-radius:999px;padding:2px 8px;margin-left:4px}
    .av-resg{margin:2px 4px 0}
    .lev-docs{display:flex;flex-wrap:wrap;align-items:center;gap:8px;padding:12px 14px;background:#f6f8fb;border-radius:12px}
    .lev-docs-label{font-size:12px;font-weight:800;letter-spacing:.6px;text-transform:uppercase;color:#8d99a9;margin-right:4px}
    .lev-docs button{border:1px solid #dbe3ee;background:#fff;color:#244c78;font-family:inherit;font-size:14px;font-weight:700;border-radius:999px;padding:7px 13px;cursor:pointer}
    .lev-docs button:hover{border-color:#9fb6d2}
    .lev-docs em{font-style:normal;font-size:11.5px;font-weight:800;color:#a56a00;background:#fff4dc;border-radius:999px;padding:2px 7px;margin-left:4px}
    .cap-areas button.done{background:#f3faf6;border-color:#cfe9da}
    .cap-areas button.done small{color:#21835a}
    .cap-tools{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:8px}
    .cap-tools button{display:flex;flex-direction:column;align-items:center;gap:2px;border:1px solid #e3e9f1;background:#fff;border-radius:14px;padding:12px 6px;font-family:inherit;cursor:pointer;color:#25364c;min-width:0}
    .cap-tools button:disabled{opacity:.5;cursor:default}
    .cap-tools span{font-size:20px;line-height:1.2}
    .cap-tools b{font-size:15px;color:#14243a}
    .cap-tools small{font-size:11.5px;color:#7d8a9b;text-align:center;line-height:1.2}
    .cap-agregado{display:flex;flex-direction:column;gap:6px;background:#f4f1fc;border:1px solid #ddd3f5;border-radius:14px;padding:14px;animation:fadeUp .25s ease both}
    .cap-agregado > b{color:#4b3590;font-size:15px}
    .cap-agregado span{font-size:15px;color:#25364c}
    .cap-agregado small{font-size:13px;color:#6d7b8e}
    .cap-agregado-btns{display:flex;gap:8px;margin-top:4px}
    .cap-agregado-btns button{flex:1}
    .cap-noenc{align-self:center;border:0;background:transparent;color:#bd4b52;font-family:inherit;font-size:14.5px;font-weight:700;padding:6px 10px;cursor:pointer}
    .busq-list{display:flex;flex-direction:column;gap:8px;max-height:46vh;overflow:auto;margin-top:6px}
    .busq-item{border:1px solid #e3e9f1;border-radius:12px;background:#fff}
    .busq-item.open{border-color:#c9d7e8;box-shadow:0 6px 16px rgba(20,36,58,.06)}
    .busq-row{display:flex;align-items:center;gap:10px;padding:10px 12px}
    .busq-row .av-txt > span{white-space:normal}
    .busq-traer{display:flex;flex-direction:column;gap:6px;padding:0 12px 12px}
    .busq-btns{display:flex;gap:8px;justify-content:flex-end;margin-top:6px}
    .busq-vacio{display:flex;flex-direction:column;align-items:flex-start;gap:8px;margin-top:8px;color:#55667c;font-size:14px}
    .busq-vacio p{margin:0}
    .av-chips button{display:flex;flex-direction:column;align-items:flex-start;gap:2px;border:1px solid #e3e9f1;background:#fff;border-radius:12px;padding:10px 12px;font-family:inherit;cursor:pointer;text-align:left;min-width:0}
    .av-chips b{font-size:20px;color:#14243a;line-height:1.1}
    .av-chips small{font-size:12.5px;color:#6d7b8e;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:100%}
    .av-chips button.on{border-color:#2d5d93;background:#eef4fb;box-shadow:0 0 0 1px #2d5d93 inset}
    .av-chips .f-pendientes b{color:#a8650f}
    .av-chips .f-fuera b{color:#bd4b52}
    .av-chips .f-revisados b{color:#21835a}
    .av-areas{display:flex;flex-direction:column;gap:8px}
    .av-area{border:1px solid #e3e9f1;background:#fff;border-radius:14px;overflow:hidden}
    .av-area.open{border-color:#c9d7e8;box-shadow:0 6px 18px rgba(20,36,58,.06)}
    .av-area-head{display:flex;align-items:center;gap:10px;width:100%;border:0;background:transparent;padding:13px 14px;font-family:inherit;text-align:left;cursor:pointer}
    .av-area-main{flex:1;min-width:0;display:flex;flex-direction:column;gap:6px}
    .av-area-main strong{font-size:15px;color:#1c2d43;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
    .av-area-count{font-size:14px;font-weight:700;color:#41536b;white-space:nowrap}
    .av-tag{font-size:12px;font-weight:700;border-radius:999px;padding:3px 9px;background:#f0f2f5;color:#6d7b8e;white-space:nowrap}
    .av-tag.ok{background:#e6f5ee;color:#21835a}
    .av-tag.warn{background:#fdecee;color:#bd4b52}
    .av-area.open .lev-chev{transform:rotate(180deg)}
    .av-items{display:flex;flex-direction:column;gap:6px;padding:0 10px 12px}
    .av-flat{padding:0}
    .av-item{border:1px solid #edf0f4;border-radius:11px;background:#fff;overflow:hidden}
    .av-item-row{display:flex;align-items:center;gap:10px;width:100%;border:0;background:transparent;padding:10px 12px;font-family:inherit;text-align:left;cursor:pointer}
    .av-ico{flex:0 0 26px;height:26px;border-radius:50%;display:grid;place-items:center;font-size:13px;font-weight:800}
    .st-ok .av-ico{background:#e6f5ee;color:#21835a}
    .st-fuera .av-ico{background:#fdecee;color:#bd4b52}
    .st-pendiente .av-ico{background:#fff4e2;color:#a8650f}
    .av-txt{flex:1;min-width:0;display:flex;flex-direction:column;gap:3px}
    .av-txt > span{font-size:14.5px;color:#25364c;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
    .av-txt small{font-size:13px;color:#7d8a9b}
    .st-fuera .av-txt small{color:#bd4b52}
    .av-go{font-size:22px;color:#9aa8b9}
    .av-detalle{display:flex;gap:12px;padding:0 12px 12px 48px;align-items:flex-start}
    .av-detalle img{width:72px;height:72px;border-radius:10px;object-fit:cover;cursor:zoom-in;background:#eef2f7}
    .av-detalle dl{margin:0;display:grid;gap:4px;font-size:13.5px}
    .av-detalle dl div{display:flex;gap:6px}
    .av-detalle dt{color:#7d8a9b;min-width:96px}
    .av-detalle dd{margin:0;color:#25364c}
    .av-pdf{margin-top:4px}
    @media (max-width:640px){
      .av-chips{grid-template-columns:repeat(3,minmax(0,1fr))}
      .cap-tools{grid-template-columns:repeat(2,minmax(0,1fr))}
      .av-chips button:first-child{grid-column:span 3;flex-direction:row;align-items:baseline;gap:8px}
      .av-area-head{flex-wrap:wrap}
      .av-area-main{flex-basis:100%}
      .av-detalle{padding-left:12px}
    }
    .btn-eliminar-lev{margin-left:auto;background:none;border:1px solid transparent;color:#a3474e;font-size:14px;font-weight:700;padding:11px 14px;border-radius:8px;cursor:pointer}
    .btn-eliminar-lev:hover{background:#fdf1f2;border-color:#efc4c7}
    .lev-more{align-self:flex-start;border:0;background:transparent;padding:4px 0;font-family:inherit;font-size:14px;font-weight:700;color:#2d5d93;cursor:pointer}
    .lev-detail dl{margin:0;display:grid;grid-template-columns:repeat(auto-fill,minmax(230px,1fr));gap:12px}
    .lev-detail dt{font-size:12px;font-weight:800;letter-spacing:.6px;color:#8d99a9;text-transform:uppercase;margin-bottom:3px}
    .lev-detail dd{margin:0;font-size:14px;color:#25364c}
    .lev-detail small{color:#8d99a9}
    .lev-person{display:block}
    .wizard{max-width:560px!important}
    .wiz-steps{display:flex;gap:6px;padding:0 24px 6px}
    .wiz-steps span{flex:1;height:5px;border-radius:4px;background:#e5eaf1;transition:background .3s}
    .wiz-steps span.on{background:#244c78}
    .wiz-people{display:flex;flex-direction:column;gap:8px}
    .wiz-people div{display:grid;grid-template-columns:1fr 1fr auto;gap:6px;align-items:center}
    .wiz-ops{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:8px}
    .wiz-ops span{display:inline-flex;align-items:center;gap:6px;font-size:14px;font-weight:600;color:#244c78;background:#e5eef8;border-radius:20px;padding:6px 8px 6px 12px}
    .wiz-ops button{border:0;background:transparent;color:#244c78;font-size:16px;cursor:pointer}
    .wiz-summary{margin-top:14px;font-size:14px;color:#41546b;background:#f1f6fc;border-radius:10px;padding:12px 14px}
    .cap{max-width:560px;margin:0 auto;display:flex;flex-direction:column;gap:14px}
    .cap-top{position:sticky;top:0;z-index:5;background:#f4f6fa;padding:4px 0 10px;display:flex;flex-direction:column;gap:6px}
    .cap-top strong{font-size:17px;color:#14243a}
    .cap-top small{font-size:13px;color:#7d8a9b}
    .cap-loading{display:flex;flex-direction:column;align-items:center;gap:12px;padding:40px 0;color:#41546b}
    .cap-q{margin:6px 0 0;font-size:22px;color:#14243a}
    .cap-hint{margin:0;font-size:14px;color:#7d8a9b}
    .cap-search,.cap-clave input{height:48px;font-size:16px!important}
    .cap-areas{display:flex;flex-direction:column;gap:8px}
    .cap-areas button,.cap-list button{display:flex;align-items:center;justify-content:space-between;gap:10px;text-align:left;border:1px solid #e3e9f1;background:#fff;border-radius:12px;padding:14px 16px;font-family:inherit;font-size:15px;color:#25364c;cursor:pointer}
    .cap-areas small{font-size:13px;font-weight:700;color:#7d8a9b;flex-shrink:0}
    .cap-list button{justify-content:flex-start}
    .cap-list span{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
    .cap-list{display:flex;flex-direction:column;gap:6px;max-height:360px;overflow:auto}
    .cap-area-chip{align-self:flex-start;border:0;background:#e5eef8;color:#244c78;font-family:inherit;font-size:14px;font-weight:700;border-radius:20px;padding:8px 14px;cursor:pointer}
    .cap-ok{background:#e8f6ee;color:#1d7a46;border-radius:12px;padding:12px 14px;font-size:14.5px;font-weight:700;animation:fadeUp .25s ease both}
    .cap-big{display:flex;align-items:center;justify-content:center;gap:12px;height:76px;border:0;border-radius:18px;background:linear-gradient(135deg,#244c78,#17304f);color:#fff;font-family:inherit;font-size:19px;font-weight:800;cursor:pointer;box-shadow:0 10px 24px #17304f33}
    .cap-big span{font-size:26px}
    .cap-big.secondary{height:56px;font-size:15.5px;background:#fff;color:#244c78;border:1.5px solid #c9d6e6;box-shadow:none}
    .cap-scanner{position:relative;display:flex;flex-direction:column;gap:10px}
    .cap-scanner video{width:100%;border-radius:16px;background:#000;aspect-ratio:3/4;object-fit:cover}
    .cap-frame{position:absolute;top:18%;left:15%;width:70%;aspect-ratio:1;border:3px solid #ffffffcc;border-radius:18px;box-shadow:0 0 0 2000px #0005;pointer-events:none}
    .cap-clave{display:flex;gap:8px}
    .cap-bien{background:#fff;border:1px solid #e3e9f1;border-radius:16px;padding:16px;display:flex;flex-direction:column;gap:6px}
    .cap-bien .chip-key{align-self:flex-start}
    .cap-bien h2{margin:4px 0 0;font-size:20px;color:#14243a;text-transform:uppercase}
    .cap-bien small{font-size:14px;color:#6f7d8f}
    .cap-warn{margin-top:6px;font-size:13.5px;font-weight:700;color:#a56a00;background:#fff4dc;border-radius:10px;padding:9px 11px}
    .cap-info{font-size:14px;color:#2d5d93;background:#eef3f9;border-radius:10px;padding:10px 12px}
    .cap-tiene-foto{display:flex;align-items:center;gap:12px}
    .cap-tiene-foto img{width:64px;height:64px;border-radius:10px;object-fit:cover;cursor:zoom-in;background:#dfe7f1}
    .drawer-foto{padding-top:14px}
    .foto-mini{display:block;position:relative;width:100%;padding:0;border:0;border-radius:14px;overflow:hidden;cursor:zoom-in;background:#eef2f7}
    .foto-mini img{display:block;width:100%;height:220px;object-fit:cover}
    .foto-mini span{position:absolute;right:10px;bottom:10px;background:rgba(15,30,50,.72);color:#fff;font-size:13px;font-weight:700;padding:6px 10px;border-radius:999px}
    .foto-visor{position:fixed;top:0;right:0;bottom:0;left:0;inset:0;z-index:3000;background:rgba(8,15,28,.88);display:flex;flex-direction:column;align-items:center;justify-content:center;gap:14px;padding:20px}
    .foto-visor img{max-width:100%;max-height:calc(100vh - 110px);border-radius:12px;box-shadow:0 20px 60px rgba(0,0,0,.5);background:#fff}
    .foto-visor-pie{display:flex;align-items:center;gap:16px;color:#fff;font-size:15px;flex-wrap:wrap;justify-content:center}
    .foto-visor-pie button{background:#fff;color:#1c2d43;border:0;border-radius:999px;padding:8px 16px;font-weight:700;font-size:14px;cursor:pointer}
    .cap-step{background:#fff;border:1px solid #e3e9f1;border-radius:16px;padding:16px;display:flex;flex-direction:column;gap:12px}
    .cap-photo{display:flex;align-items:center;justify-content:center;gap:10px;height:110px;border:2px dashed #b9cbe0;border-radius:14px;color:#244c78;font-size:17px;font-weight:800;cursor:pointer;background:#f6f9fc}
    .cap-photo span{font-size:28px}
    .cap-preview{display:flex;flex-direction:column;align-items:center;gap:10px}
    .cap-preview img{width:100%;max-height:320px;object-fit:contain;border-radius:12px;background:#f1f4f8}
    .cap-options{display:grid;grid-template-columns:1fr 1fr;gap:10px}
    .cap-options button{height:58px;border:1.5px solid #d9e2ee;background:#fff;border-radius:14px;font-family:inherit;font-size:16.5px;font-weight:700;color:#25364c;cursor:pointer;transition:.15s}
    .cap-options button.on{border-color:#244c78;background:#244c78;color:#fff;box-shadow:0 6px 16px #244c7840}
    .cap-footer{display:flex;gap:10px;position:sticky;bottom:0;background:#f4f6fa;padding:10px 0}
    .cap-footer button{flex:1;height:54px;font-size:16px}
    .t-actas{background:linear-gradient(140deg,#7a4a00,#d98a00)}
    .doc-badge.acta{background:#d98a00}
    .doc-row-ico.acta{background:#fff4dc;color:#a56a00}
    @media(max-width:700px){.lev-stats{grid-template-columns:repeat(2,minmax(0,1fr))}.lev-head{flex-wrap:wrap}.wiz-people div{grid-template-columns:1fr}}

    /* ---------- DOCUMENTOS ---------- */
    .t-alta{background:linear-gradient(140deg,#0f4c3a,#21835a)}
    .t-baja{background:linear-gradient(140deg,#6b1f2a,#c2414b)}
    .t-transfer{background:linear-gradient(140deg,#17304f,#2d5d93)}
    .t-resguardo{background:linear-gradient(140deg,#3b2a6b,#6a4fb3)}
    .inv-tile.disabled{cursor:default;opacity:.55;filter:saturate(.55)}
    .inv-tile.disabled:hover{transform:none;box-shadow:0 10px 30px #17263d26}
    .tile-soon{align-self:flex-start;font-size:13px;font-weight:700;background:#ffffff26;border:1px solid #ffffff38;border-radius:20px;padding:6px 14px}
    .doc-list{display:flex;flex-direction:column;padding:4px 12px 14px;gap:8px}
    .doc-row{display:flex;align-items:center;gap:14px;text-align:left;font-family:inherit;background:#fff;border:1px solid #e5eaf1;border-radius:14px;padding:13px 16px;cursor:pointer;transition:transform .15s,box-shadow .15s,border-color .15s}
    .doc-row:hover:not(:disabled){transform:translateY(-1px);border-color:#c9d6e6;box-shadow:0 8px 20px #233c5712}
    .doc-row.anulado{opacity:.6}
    .doc-row-ico{width:44px;height:44px;border-radius:12px;background:#e8f6ee;color:#21835a;display:grid;place-items:center;flex-shrink:0}
    .doc-row-ico.baja{background:#fdeced;color:#c2414b}
    .doc-row-ico.resg{background:#f1edfb;color:#6a4fb3}
    .doc-badge.resg{background:#6a4fb3}
    .doc-estado{display:inline-block;font-size:11px;font-weight:800;letter-spacing:.5px;color:#55667c;background:#eef1f5;border-radius:999px;padding:3px 9px;margin-left:6px;vertical-align:middle}
    .doc-estado.borrador{color:#a56a00;background:#fff4dc}
    .doc-row-ico svg{width:22px;height:22px}
    .doc-row-main{flex:1;min-width:0;display:flex;flex-direction:column;gap:3px}
    .doc-row-top{display:flex;align-items:center;gap:8px}
    .doc-row-main strong{font-size:15px;color:#14243a;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
    .doc-row-main small{font-size:13px;color:#7d8a9b;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
    .doc-row-date{text-align:right;font-size:14px;font-weight:700;color:#41546b;flex-shrink:0}
    .doc-row-date small{display:block;font-size:12px;font-weight:500;color:#8d99a9}
    .doc-row-go{font-size:14px;font-weight:700;color:#2d5d93;flex-shrink:0;min-width:44px;text-align:right}
    .doc-anulado{display:inline-block;font-size:11px;font-weight:800;letter-spacing:1px;color:#fff;background:#c2414b;border-radius:6px;padding:3px 8px;vertical-align:middle}
    @media(max-width:700px){.doc-row{flex-wrap:wrap}.doc-row-date{text-align:left}.doc-row-go{display:none}}

    /* ---------- SALUDO DE INICIO ---------- */
    .welcome{min-height:calc(100vh - 150px);background:#fff;border:1px solid #e7ecf2;border-radius:16px;display:flex;flex-direction:column;align-items:center;padding:32px 20px 22px;box-shadow:0 3px 12px #233c5708}
    .welcome-inner{flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;max-width:520px;width:100%}
    .welcome-logo{width:min(300px,78%);height:auto;margin-bottom:34px;animation:fadeUp .6s ease both}
    .push-invita{display:flex;align-items:center;gap:12px;margin:22px auto 0;max-width:460px;text-align:left;background:#fff;border:1px solid #dfe7f1;border-radius:16px;padding:14px 14px 14px 16px;box-shadow:0 8px 24px #17263d12;animation:fadeUp .5s .2s ease both}
    .push-ico{font-size:24px}
    .push-txt{flex:1;min-width:0;display:flex;flex-direction:column;gap:2px}
    .push-txt strong{font-size:15px;color:#14243a}
    .push-txt small{font-size:13px;color:#6d7b8e}
    .push-x{border:0;background:transparent;color:#9aa6b6;font-size:20px;cursor:pointer;padding:0 4px}
    .push-instalar{margin-top:12px;border:1px solid #d6e0ec;background:#fff;color:#244c78;font-family:inherit;font-size:14px;font-weight:700;border-radius:999px;padding:9px 16px;cursor:pointer}
    .etq-stats{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:8px;margin-bottom:10px}
    .etq-stats div{background:#f6f8fb;border-radius:12px;padding:10px 12px}
    .etq-stats b{display:block;font-size:20px;color:#14243a}
    .etq-stats small{font-size:12.5px;color:#7d8a9b}
    .etq-sedes{display:flex;flex-direction:column;gap:4px;margin:4px 0 10px}
    .etq-sedes div{display:grid;grid-template-columns:1fr auto 90px;gap:10px;font-size:14px;padding:6px 10px;background:#f9fafc;border-radius:8px}
    .etq-sedes small{color:#7d8a9b;text-align:right}
    .etq-hojas{display:flex;gap:8px;align-items:center;margin-bottom:6px}
    .etq-hojas input{width:90px;height:44px;text-align:center;font-size:18px!important;font-weight:800}
    .etq-hojas button{width:44px;height:44px;font-size:20px;padding:0}
    .etq-btns{display:flex;gap:8px;flex-wrap:wrap}
    .etq-vacia{font-size:42px;opacity:.35}
    .etq-foto{display:inline-flex;align-items:center;justify-content:center;width:52px;cursor:pointer}
    .etq-aviso{background:#f1edfb!important;border-color:#ddd3f5!important;color:#4b3590!important}
    .cm-aviso{display:flex;align-items:center;gap:12px;margin:12px auto 0;max-width:460px;text-align:left;background:#fff;border:1px solid #dfe7f1;border-radius:16px;padding:12px 14px 12px 16px;animation:fadeUp .5s .25s ease both}
    .cm-n{font-weight:800}
    .cm-n.vencida{color:#bd4b52}
    .cm-n.proxima{color:#a8650f}
    .cm-form{display:flex;flex-direction:column;gap:10px;margin:0 0 16px;padding:12px 14px;background:#f6f8fb;border-radius:12px}
    .cm-row{display:flex;align-items:center;gap:12px;flex-wrap:wrap;min-height:40px}
    .cm-check{display:flex!important;align-items:center;gap:8px;margin:0!important;font-size:14px!important;color:#25364c!important;cursor:pointer;min-width:220px}
    .cm-check input{width:18px;height:18px}
    .cm-row input[type=date]{height:40px;max-width:200px}
    .cm-chips{display:flex;flex-direction:column;gap:8px}
    .cm-chip{align-self:flex-start;font-size:14px;font-weight:700;border-radius:999px;padding:7px 13px;background:#eef1f5;color:#41546b}
    .cm-chip.cm-vencida{background:#fdecee;color:#bd4b52}
    .cm-chip.cm-proxima{background:#fff4dc;color:#a8650f}
    .cm-chip.cm-ok{background:#e8f6ee;color:#21835a}
    .welcome-chip{font-size:13.5px;font-weight:700;color:#4f6f95;background:#eef3f9;padding:6px 14px;border-radius:20px;margin-bottom:18px;animation:fadeUp .6s .1s ease both}
    .welcome h1{margin:0 0 14px;font-size:36px;font-weight:800;letter-spacing:-.6px;color:#111827;animation:fadeUp .6s .18s ease both}
    .welcome-inner p{margin:0;font-size:16.5px;line-height:1.7;color:#4b5563;animation:fadeUp .6s .26s ease both}
    .welcome-foot{margin-top:28px;font-size:12.5px;font-weight:600;letter-spacing:2.2px;color:#9aa3af;text-align:center}
    @media(max-width:700px){.welcome{min-height:calc(100vh - 120px)}.welcome h1{font-size:29px}.welcome-inner p{font-size:15.5px}}

    /* ---------- LOGIN ---------- */
    .lg-screen{min-height:100vh;background:#fff;display:flex;justify-content:center;padding:48px 16px 32px}
    .lg-card{width:min(440px,100%);display:flex;flex-direction:column;align-items:stretch;animation:fadeUp .5s ease both}
    .lg-logo-img{display:block;width:min(340px,86%);height:auto;margin:24px auto 44px}
    .lg-titles{text-align:center;margin-bottom:40px}
    .lg-titles h1{margin:0 0 12px;font-size:23px;font-weight:800;letter-spacing:6px;color:#111827;padding-left:6px}
    .lg-titles p{margin:0;font-size:14.5px;font-weight:700;letter-spacing:5px;color:#5f6f86;padding-left:5px}
    .lg-field{position:relative;margin-bottom:18px}
    .lg-field input{width:100%;height:56px;padding:0 52px 0 22px;border:1.5px solid #dfe3ec;border-radius:14px;background:#f8fafc;font-family:"SFMono-Regular",Consolas,"Courier New",monospace;font-size:16.5px;letter-spacing:1px;color:#111827;transition:border-color .2s,box-shadow .2s,background .2s}
    .lg-field input::placeholder{color:#9aa3af;font-weight:700;letter-spacing:2px}
    .lg-field input:focus{outline:none;border-color:#111827;background:#fff;box-shadow:0 0 0 4px #11182714}
    .lg-field.invalid input{border-color:#e2a1a6}
    .lg-eye{position:absolute;right:9px;top:50%;transform:translateY(-50%);width:38px;height:38px;border:0;background:transparent;border-radius:10px;display:grid;place-items:center;cursor:pointer}
    .lg-eye:hover{background:#eef1f5}
    .lg-eye svg{width:18px;height:18px;fill:none;stroke:#7a8494;stroke-width:2;stroke-linecap:round;stroke-linejoin:round}
    .lg-error{display:flex;gap:9px;align-items:flex-start;background:#fdf0f1;color:#b4404a;border:1px solid #f5d3d6;border-radius:12px;padding:11px 14px;font-size:14.5px;line-height:1.45;margin:-2px 0 16px;animation:fadeUp .25s ease both}
    .lg-error svg{width:16px;height:16px;flex-shrink:0;fill:none;stroke:currentColor;stroke-width:2;stroke-linecap:round}
    .lg-submit{height:56px;margin-top:6px;border:0;border-radius:14px;background:#111827;color:#fff;font-family:inherit;font-size:16.5px;font-weight:800;letter-spacing:4px;display:flex;align-items:center;justify-content:center;gap:10px;cursor:pointer;transition:background .2s,transform .15s,box-shadow .2s}
    .lg-submit:hover:not(:disabled){background:#1f2a3d;transform:translateY(-1px);box-shadow:0 10px 24px #11182733}
    .lg-submit:disabled{opacity:.6;cursor:not-allowed}
    .lg-forgot{align-self:center;margin-top:26px;border:0;background:transparent;font-family:inherit;font-size:15.5px;color:#5f6f86;text-decoration:underline;text-underline-offset:3px;cursor:pointer}
    .lg-forgot:hover{color:#111827}
    .lg-recover{text-align:center;margin:-12px 0 24px}
    .lg-recover h2,.lg-sent h2{margin:0 0 10px;font-size:19px;color:#111827}
    .lg-recover p,.lg-sent p{margin:0;font-size:15px;line-height:1.6;color:#5f6f86}
    .lg-rules{list-style:none;display:flex;flex-wrap:wrap;gap:6px 16px;margin:-4px 0 18px;padding:0 4px;font-size:14px;color:#9aa3af}
    .lg-rules li::before{content:'○ '}
    .lg-rules li.ok{color:#21835a}
    .lg-rules li.ok::before{content:'✓ '}
    .lg-sent{text-align:center;margin-top:-8px;animation:fadeUp .3s ease both}
    .lg-sent-ico{width:54px;height:54px;border-radius:50%;background:#e8f6ee;display:grid;place-items:center;margin:0 auto 16px}
    .lg-sent-ico svg{width:26px;height:26px;fill:none;stroke:#21835a;stroke-width:2.5;stroke-linecap:round;stroke-linejoin:round}
    .lg-sent ol{text-align:left;margin:20px 0 0;padding:16px 18px 16px 38px;background:#f6f8fb;border-radius:12px;font-size:15px;line-height:1.9;color:#374151}
    .rec-panel{background:#fff;border:1px solid #f0dcae;border-radius:12px;margin-bottom:20px;overflow:hidden;box-shadow:0 3px 12px #233c5708}
    .rec-head{display:flex;align-items:center;gap:12px;padding:15px 20px;background:#fffaf0;border-bottom:1px solid #f5e6c4}
    .rec-head-ico{width:36px;height:36px;border-radius:10px;background:#fdecc8;color:#a56a00;display:grid;place-items:center;font-size:19px;flex-shrink:0}
    .rec-head strong{display:block;font-size:15.5px;color:#25364c}
    .rec-head small{display:block;font-size:13.5px;color:#8a7a55;margin-top:2px}
    .rec-row{display:flex;align-items:center;gap:14px;padding:13px 20px;border-bottom:1px solid #f1f4f8}
    .rec-row:last-child{border-bottom:0}
    .rec-avatar{width:36px;height:36px;border-radius:50%;background:#eef3f9;color:#244c78;display:grid;place-items:center;font-size:14px;font-weight:800;flex-shrink:0}
    .rec-info{flex:1;min-width:0}
    .rec-info strong{display:block;font-size:15px;color:#25364c;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .rec-info small{display:block;font-size:13.5px;color:#8c98a8;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .rec-actions{display:flex;gap:8px;flex-shrink:0}
    .temp-card{text-align:center;padding:30px 28px 24px}
    .temp-ico{width:56px;height:56px;border-radius:50%;background:#e8f6ee;color:#21835a;display:grid;place-items:center;font-size:26px;margin:0 auto 14px}
    .temp-card h2{margin:0 0 6px;font-size:21px;color:#1c2e45}
    .temp-card>p{margin:0 0 20px;font-size:15px;color:#7a889a}
    .temp-pass{display:flex;align-items:center;justify-content:center;gap:12px;background:#f4f7fb;border:1.5px dashed #c5d3e4;border-radius:12px;padding:16px}
    .temp-pass code{font-family:"SFMono-Regular",Consolas,"Courier New",monospace;font-size:28px;font-weight:800;letter-spacing:3px;color:#14243a}
    .temp-steps{text-align:left;margin:18px 0 0;padding:14px 16px 14px 34px;background:#fffaf0;border-radius:10px;font-size:14.5px;line-height:1.8;color:#5b4a26}
    @media(max-width:700px){.rec-row{flex-wrap:wrap}.rec-actions{width:100%}.rec-actions button{flex:1}}
    .lg-forgot-help{margin:12px 0 0;text-align:center;font-size:14.5px;line-height:1.55;color:#6b7686;background:#f6f8fb;border-radius:12px;padding:12px 16px;animation:fadeUp .2s ease both}
    .lg-footer{margin-top:36px;padding-top:26px;border-top:1px solid #edf0f4;text-align:center}
    .lg-footer p{margin:0 0 8px;font-size:12.5px;font-weight:600;letter-spacing:2.2px;color:#8a95a6}
    .lg-footer b{color:#1f2937;font-weight:800}
    @media (max-width:480px){.lg-screen{padding-top:24px}.lg-logo-img{margin:20px auto 34px}.lg-titles h1{letter-spacing:5px}.lg-titles p{letter-spacing:3.5px}}
    @keyframes shimmer{0%{background-position:-400px 0}100%{background-position:400px 0}}
    @keyframes spin{to{transform:rotate(360deg)}}
    @keyframes fadeUp{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:none}}
    @keyframes zoomIn{from{opacity:0;transform:scale(.96) translateY(8px)}to{opacity:1;transform:none}}
    @keyframes fadeIn{from{opacity:0}to{opacity:1}}
    .fade-in{animation:fadeUp .45s ease both}
    .sk{display:block;border-radius:6px;background:linear-gradient(90deg,#eef2f7 0%,#f7f9fc 40%,#eef2f7 80%);background-size:800px 100%;animation:shimmer 1.4s linear infinite}
    .skeleton-card{display:flex;flex-direction:column;gap:12px}
    .sk-line{height:10px}
    .sk-number{height:24px;width:45%}
    .w10{width:10%}.w15{width:15%}.w20{width:20%}.w30{width:30%}.w40{width:40%}.w60{width:60%}
    .sk-chip{height:64px;flex:1;min-width:120px;border-radius:10px}
    .sk-chips{display:flex;gap:10px;padding:0 22px 18px;flex-wrap:wrap}
    .sk-row{display:flex;align-items:center;gap:22px;padding:15px 22px;border-top:1px solid #f0f3f7}
    .sk-pill{height:20px;width:70px;border-radius:20px;margin-left:auto}
    .loading-banner{display:flex;align-items:center;gap:16px;padding:22px}
    .loading-banner strong{font-size:15.5px;color:#25364c}
    .loading-banner p{margin:5px 0 0;font-size:14px;color:#8491a3}
    .loader-ring{width:36px;height:36px;border-radius:50%;border:3px solid #dfe8f3;border-top-color:#244c78;animation:spin .9s linear infinite;flex-shrink:0}
    .mini-ring{display:inline-block;width:12px;height:12px;border-radius:50%;border:2px solid #dfe8f3;border-top-color:#244c78;animation:spin .9s linear infinite;vertical-align:-2px}
    .refreshing{font-size:13px;color:#6b7b90;display:inline-flex;align-items:center;gap:7px}

    /* ===== INVENTARIO: RESUMEN POR CLASIFICACIÓN ===== */
    .overview-panel{background:#fff;border:1px solid #e7ecf2;border-radius:12px;padding:20px 22px;margin-bottom:20px;box-shadow:0 3px 12px #233c5708}
    .overview-head{display:flex;justify-content:space-between;align-items:flex-start;gap:12px;margin-bottom:15px}
    .overview-head h2{font-size:16.5px;margin:0 0 5px;color:#25364c}
    .overview-head p{font-size:13px;color:#929eae;margin:0}
    .overview-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(170px,1fr));gap:10px}
    .overview-card{position:relative;text-align:left;font-family:inherit;border:1px solid #e8edf3;background:#fbfcfe;border-radius:10px;padding:13px 14px;cursor:pointer;display:flex;flex-direction:column;gap:6px;transition:.2s}
    .overview-card:hover{border-color:#c9d6e6;background:#fff;transform:translateY(-1px);box-shadow:0 4px 12px #233c570d}
    .overview-card.selected{border-color:#5e8fc7;background:#fff;box-shadow:0 0 0 3px #5e8fc71f}
    .overview-card.dimmed{opacity:.55}
    .overview-dot{position:absolute;top:15px;right:14px;width:8px;height:8px;border-radius:50%}
    .overview-name{font-size:12px;font-weight:800;letter-spacing:.4px;color:#6a7b90;padding-right:16px;text-transform:uppercase;line-height:1.35}
    .overview-card strong{font-size:24px;color:#1c2e45;letter-spacing:-.6px}
    .overview-bar{height:5px;background:#edf1f6;border-radius:4px;overflow:hidden}
    .overview-bar span{display:block;height:100%;border-radius:4px}
    .overview-card small{font-size:12px;color:#97a3b3}
    .link-button.inline{display:inline;margin:0;white-space:nowrap}

    /* ===== INVENTARIO: TABLA ===== */
    .inventory-filters{flex-wrap:wrap}
    .inventory-filters select{width:auto;min-width:160px;height:39px}
    .search-box{position:relative}
    .clear-search{background:transparent;border:0;color:#8a97a8;font-size:20px;cursor:pointer;padding:0 2px}
    .inventory-table th.sortable{cursor:pointer;user-select:none}
    .inventory-table th.sortable:hover{color:#2f4d70}
    .sort-icon{font-size:10.5px;margin-left:4px;color:#c2ccd8}
    .sort-icon.active{color:#244c78}
    .inventory-table .num{text-align:center}
    .chevron-col{width:28px;color:#b5c0cd;font-size:17.5px;text-align:center}
    .clickable-row:hover .chevron-col{color:#244c78}
    .row-selected{background:#f1f6fc!important}
    .page-size{width:auto;height:30px;padding:4px 8px;font-size:13px}
    .column-picker{position:relative}
    .picker-backdrop{position:fixed;top:0;right:0;bottom:0;left:0;inset:0;z-index:40}
    .picker-menu{position:absolute;right:0;top:calc(100% + 6px);z-index:41;width:270px;background:#fff;border:1px solid #e1e7ef;border-radius:12px;box-shadow:0 14px 40px #15243b2e;animation:fadeUp .18s ease both}
    .picker-head{padding:14px 16px 10px;border-bottom:1px solid #eef1f5}
    .picker-head strong{display:block;font-size:14px;color:#25364c}
    .picker-head small{font-size:12px;color:#8b97a7}
    .picker-presets{display:flex;gap:6px;padding:10px 16px}
    .picker-presets button{flex:1;border:1px solid #dce3ec;background:#f7f9fc;border-radius:7px;padding:6px 4px;font-size:12px;font-weight:700;color:#4d6077;cursor:pointer}
    .picker-presets button:hover{background:#eaf1f9;border-color:#b7c6d9}
    .picker-list{max-height:290px;overflow:auto;padding:4px 8px 10px}
    .picker-list label{display:flex;align-items:center;gap:9px;padding:7px 8px;border-radius:6px;font-size:14px;color:#3d4f66;cursor:pointer}
    .picker-list label:hover{background:#f5f8fb}
    .picker-list label.locked{color:#97a3b3;cursor:default}
    .picker-list input{width:auto;margin:0}
    .picker-list small{margin-left:auto;font-size:11.5px;color:#aab5c3;text-transform:uppercase;letter-spacing:.5px}

    .filters-menu{width:320px;max-height:calc(100vh - 160px);display:flex;flex-direction:column;overflow:auto}
    .filters-active{border-color:#5e8fc7!important;color:#244c78!important;background:#f5f9fe!important}
    .filter-fields{display:grid;gap:10px;padding:12px 16px 14px}
    .filter-fields label{display:flex;flex-direction:column;gap:5px;font-size:12px;font-weight:700;color:#64748a}
    .filter-fields select{height:36px;font-weight:400}
    .picker-subhead{border-top:1px solid #eef1f5}
    .picker-footer{position:sticky;bottom:0;display:flex;justify-content:space-between;align-items:center;gap:10px;padding:12px 16px;border-top:1px solid #eef1f5;background:#fafbfd}
    .picker-footer .link-button:disabled{opacity:.4;cursor:default;text-decoration:none}
    .active-filters{display:flex;flex-wrap:wrap;gap:6px;align-items:center}
    .filter-chip{display:inline-flex;align-items:center;gap:5px;background:#eaf1f9;color:#2f5580;border:1px solid #d3e1f1;border-radius:20px;padding:5px 6px 5px 11px;font-size:13px;animation:fadeUp .15s ease both}
    .filter-chip b{font-weight:700}
    .filter-chip button{border:0;background:#d3e1f1;color:#2f5580;width:18px;height:18px;border-radius:50%;cursor:pointer;font-size:15px;line-height:1;display:grid;place-items:center;padding:0}
    .filter-chip button:hover{background:#244c78;color:#fff}

    /* ===== INVENTARIO: PANEL LATERAL ===== */
    .drawer-backdrop{position:fixed;top:0;right:0;bottom:0;left:0;inset:0;background:#15243b55;z-index:90;animation:fadeIn .2s ease both}
    .drawer{position:fixed;top:0;right:0;bottom:0;left:0;inset:0;margin:auto;width:min(780px,calc(100% - 32px));height:fit-content;max-height:calc(100vh - 48px);max-height:calc(100dvh - 48px);background:#fff;z-index:91;display:flex;flex-direction:column;border-radius:16px;overflow:hidden;box-shadow:0 24px 80px #15243b40;animation:zoomIn .22s ease both}
    .drawer-section dl{grid-template-columns:repeat(3,minmax(0,1fr))!important}
    .drawer-header{display:flex;justify-content:space-between;gap:12px;padding:24px 24px 18px;border-bottom:1px solid #edf0f4}
    .drawer-header h2{margin:6px 0 10px;font-size:21px;color:#1c2d43;line-height:1.3}
    .drawer-badges{display:flex;gap:6px;flex-wrap:wrap}
    .drawer-body{flex:1;overflow:auto;padding:6px 24px 20px}
    .drawer-section{padding:16px 0;border-bottom:1px solid #f0f3f7}
    .drawer-section:last-of-type{border-bottom:0}
    .drawer-section h3{font-size:12px;font-weight:800;letter-spacing:1px;color:#6482a6;margin:0 0 12px;text-transform:uppercase}
    .drawer-section dl{margin:0;display:grid;grid-template-columns:1fr 1fr;gap:14px 18px}
    .drawer-section dl div{min-width:0}
    .drawer-section dt{font-size:12px;color:#8996a7;margin-bottom:4px}
    .drawer-section dd{margin:0;font-size:14px;font-weight:600;color:#2f4359;overflow-wrap:anywhere}
    .drawer-footer{display:flex;gap:10px;justify-content:flex-end;padding:16px 24px;border-top:1px solid #edf0f4;background:#fafbfd}
    .drawer .info-box{margin-top:14px}
    @media(max-width:700px){.overview-grid{grid-template-columns:repeat(2,minmax(0,1fr))}.inventory-filters select{flex:1;min-width:0}.picker-menu{right:auto;left:0}.drawer{width:calc(100% - 20px);max-height:calc(100vh - 20px);max-height:calc(100dvh - 20px)}.drawer-section dl{grid-template-columns:1fr 1fr!important}.drawer-footer button{flex:1}.sk-row span:nth-child(n+4){display:none}}

    /* ===== DIÁLOGO DE CONFIRMACIÓN ===== */
    .confirm-backdrop{position:fixed;top:0;right:0;bottom:0;left:0;inset:0;background:#15243b80;z-index:200;display:flex;align-items:center;justify-content:center;padding:16px;animation:fadeIn .15s ease both;backdrop-filter:blur(2px)}
    .confirm-card{width:min(460px,100%);max-height:calc(100vh - 32px);overflow:auto;background:#fff;border-radius:18px;padding:28px 26px 22px;text-align:center;box-shadow:0 30px 90px #0b172845;animation:zoomIn .2s ease both}
    .confirm-icon{width:58px;height:58px;margin:0 auto 16px;border-radius:50%;display:grid;place-items:center;font-size:28px}
    .confirm-peligro .confirm-icon{background:#fdeced;color:#c0474e;box-shadow:0 0 0 8px #fdeced80}
    .confirm-confirmar .confirm-icon{background:#e9f1fc;color:#2f5f97;box-shadow:0 0 0 8px #e9f1fc80}
    .confirm-card h2{margin:0 0 8px;font-size:21px;color:#1c2d43;letter-spacing:-.3px}
    .confirm-card>p{margin:0 auto;max-width:360px;font-size:15px;color:#6b7b90;line-height:1.55}
    .confirm-asset{margin:18px 0 0;background:#f6f8fb;border:1px solid #e8edf3;border-radius:12px;padding:12px 14px;display:flex;flex-direction:column;gap:4px;align-items:center}
    .confirm-asset strong{font-size:15px;color:#2b3d53;text-transform:uppercase}
    .confirm-asset small{font-size:13px;color:#8b97a7}
    .confirm-changes{margin:16px 0 0;text-align:left;border:1px solid #e8edf3;border-radius:12px;overflow:hidden}
    .confirm-changes-title{font-size:12px;font-weight:800;letter-spacing:1px;color:#6482a6;text-transform:uppercase;padding:10px 14px;background:#f7f9fc;border-bottom:1px solid #edf0f4}
    .change-row{display:grid;grid-template-columns:110px minmax(0,1fr) 16px minmax(0,1fr);gap:8px;align-items:center;padding:9px 14px;border-bottom:1px solid #f0f3f7;font-size:13px}
    .change-row:last-child{border-bottom:0}
    .change-label{font-weight:700;color:#5b6c82}
    .change-old{color:#9aa6b6;text-decoration:line-through;overflow-wrap:anywhere}
    .change-arrow{color:#b5c0cd;text-align:center}
    .change-new{color:#8a5a12;font-weight:700;overflow-wrap:anywhere}
    .confirm-note{margin-top:14px;background:#fff8ea;border:1px solid #f3e1b8;color:#8a5a12;border-radius:10px;padding:10px 12px;font-size:13px;line-height:1.45;text-align:left}
    .confirm-card .inline-error{margin:14px 0 0;text-align:left}
    .confirm-actions{display:flex;gap:10px;margin-top:22px}
    .confirm-actions button{flex:1;padding:12px 14px}
    .btn-danger-solid{border-radius:8px;padding:11px 16px;font-size:14px;font-weight:700;cursor:pointer;display:inline-flex;align-items:center;justify-content:center;gap:8px;background:#c0474e;color:#fff;border:1px solid #c0474e;transition:.2s}
    .btn-danger-solid:hover{background:#a63a40}
    .btn-danger-solid:disabled{opacity:.6;cursor:not-allowed}
    .mini-ring.light{border-color:#ffffff55;border-top-color:#fff}
    @media(max-width:520px){.change-row{grid-template-columns:1fr;gap:2px}.change-arrow{display:none}.confirm-actions{flex-direction:column-reverse}}

    /* ===== ETIQUETAS QR ===== */
    .qr-block{display:flex;gap:16px;align-items:center}
    .qr-frame{width:118px;height:118px;flex-shrink:0;border:1px solid #e3e9f0;border-radius:10px;display:grid;place-items:center;background:#fff;padding:6px}
    .qr-frame img{width:100%;height:100%;image-rendering:pixelated}
    .qr-info{display:flex;flex-direction:column;align-items:flex-start;gap:4px;min-width:0}
    .qr-info small{font-size:12px;color:#8996a7}
    .qr-id{font-family:'Courier New',monospace;font-size:16.5px;letter-spacing:1.5px;color:#244c78}
    .qr-info p{font-size:13px;color:#7d8b9d;margin:2px 0 8px;line-height:1.45}
    .label-config{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1fr);gap:24px;padding:20px 24px}
    .label-config h3{font-size:12px;font-weight:800;letter-spacing:1px;color:#6482a6;margin:0 0 10px;text-transform:uppercase}
    .label-options h3:not(:first-child){margin-top:20px}
    .option-card{display:flex;gap:11px;align-items:flex-start;border:1px solid #e1e7ef;border-radius:10px;padding:12px 13px;margin-bottom:8px;cursor:pointer;transition:.15s}
    .option-card:hover{border-color:#b7c6d9}
    .option-card.selected{border-color:#5e8fc7;background:#f5f9fe;box-shadow:0 0 0 3px #5e8fc71a}
    .option-card input{width:auto;margin:2px 0 0}
    .option-card strong{display:block;font-size:14px;color:#25364c;margin-bottom:3px}
    .option-card small{font-size:13px;color:#8491a3;line-height:1.4}
    .label-options .warning-box{margin-top:12px}
    .label-preview-wrap{background:#f4f6fa;border-radius:12px;padding:16px;display:flex;flex-direction:column}
    .label-preview{background:#fff;border:1px solid #c9d3df;border-radius:9px;padding:14px;display:flex;gap:14px;align-items:center;box-shadow:0 6px 18px #233c5714;margin:auto 0;width:100%;max-width:340px;align-self:center}
    .label-preview.compacta{max-width:240px;padding:10px;gap:10px}
    .lp-qr{width:112px;height:112px;flex-shrink:0;image-rendering:pixelated}
    .label-preview.compacta .lp-qr{width:76px;height:76px}
    .lp-info{display:flex;flex-direction:column;gap:4px;min-width:0}
    .lp-brand{display:flex;align-items:center;gap:5px;font-size:11.5px;font-weight:800;letter-spacing:1px;color:#244c78}
    .lp-brand span{width:15px;height:15px;border-radius:4px;background:#17263d;color:#fff;display:grid;place-items:center;font-size:11.5px}
    .lp-clave{font-size:19px;font-weight:800;color:#17263d}
    .label-preview.compacta .lp-clave{font-size:15px}
    .lp-articulo{font-size:12px;font-weight:700;color:#33465e;text-transform:uppercase;line-height:1.25;overflow:hidden;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical}
    .lp-meta{font-size:11.5px;color:#5d6f86}
    .lp-id{font-family:'Courier New',monospace;font-size:11.5px;color:#5d6f86;letter-spacing:.8px}
    .label-preview-wrap .preview-note{margin-top:14px;line-height:1.5}
    @media(max-width:800px){.label-config{grid-template-columns:1fr}}

    /* ===== CATÁLOGOS ===== */
    .catalog-layout{display:grid;grid-template-columns:230px minmax(0,1fr);gap:18px;align-items:start}
    .catalog-nav{background:#fff;border:1px solid #e7ecf2;border-radius:12px;padding:12px;box-shadow:0 3px 12px #233c5708;position:sticky;top:16px}
    .catalog-group{font-size:11.5px;font-weight:800;letter-spacing:1.2px;color:#8a9ab0;margin:12px 8px 6px}
    .catalog-group:first-child{margin-top:4px}
    .catalog-nav button{width:100%;display:flex;align-items:center;justify-content:space-between;gap:8px;border:0;background:transparent;padding:9px 10px;border-radius:7px;font-size:14px;color:#4d6077;cursor:pointer;text-align:left}
    .catalog-nav button:hover{background:#f3f6fa}
    .catalog-nav button.selected{background:#eaf1f9;color:#244c78;font-weight:700}
    .catalog-nav button.selected .tab-count{background:#244c78;color:#fff}
    .catalog-panel{min-width:0}
    .catalog-main{color:#2b3d53;font-size:13px}
    .inactive-row td{opacity:.6}
    .inline-check{display:inline-flex!important;align-items:center;gap:8px;font-size:14px!important;font-weight:600!important;color:#4d6077!important;white-space:nowrap;cursor:pointer;margin:0!important}
    .inline-check input{width:auto;margin:0}
    .full-span{grid-column:1/-1}
    .pager{display:flex;align-items:center;gap:10px}
    .modal-card .form-grid select,.modal-card .form-grid input{height:40px}
    @media(max-width:900px){.catalog-layout{grid-template-columns:1fr}.catalog-nav{position:static;display:flex;flex-wrap:wrap;gap:4px}.catalog-group{width:100%;margin:8px 4px 2px}.catalog-nav button{width:auto}}
    @media(max-width:700px){.table-footer{flex-direction:column;gap:10px;align-items:flex-start}}

    /* ===== CONTRASEÑA ===== */
    .sidebar-bottom .logout{padding:0 2px}
    .password-card>p{line-height:1.5}
    .password-modal{max-width:460px}
    .password-form label{display:block;font-size:12px;font-weight:700;color:#64748a;margin:16px 0 7px}
    .password-form label:first-child{margin-top:0}
    .password-form input{height:42px}
    .password-rules{list-style:none;padding:0;margin:16px 0 0;display:grid;gap:6px}
    .password-rules li{font-size:13px;color:#97a3b3;display:flex;align-items:center;gap:8px}
    .password-rules li::before{content:'○';font-size:14px}
    .password-rules li.ok{color:#21835a}
    .password-rules li.ok::before{content:'✓';font-weight:800}
    .link-button{display:block;margin:14px auto 0;background:transparent;border:0;color:#5f7590;font-size:14px;cursor:pointer;text-decoration:underline}

    /* ===== BANDEJAS DE SOLICITUDES ===== */
    .nav-badge{margin-left:auto;background:#e8913a;color:#fff;font-size:12px;font-weight:800;min-width:20px;height:20px;padding:0 6px;border-radius:10px;display:inline-grid;place-items:center}
    .red{background:#fdeced;color:#bd5056}
    .stat-button{font-family:inherit;text-align:left;cursor:pointer;width:100%;transition:.2s}
    .stat-button:hover{border-color:#c9d6e6;transform:translateY(-1px)}
    .stat-button.stat-selected{border-color:#5e8fc7;box-shadow:0 0 0 3px #5e8fc71f}
    .segmented{display:inline-flex;background:#eef2f7;border-radius:9px;padding:3px}
    .segmented button{border:0;background:transparent;padding:8px 13px;border-radius:7px;font-size:14px;font-weight:700;color:#647890;cursor:pointer}
    .segmented button.selected{background:#fff;color:#244c78;box-shadow:0 1px 4px #1f3b5a1f}
    .request-tabs{display:flex;gap:4px;padding:14px 22px 0;border-bottom:1px solid #edf0f4;flex-wrap:wrap}
    .request-tabs button{border:0;background:transparent;padding:11px 14px;font-size:14px;font-weight:700;color:#7a8a9e;cursor:pointer;border-bottom:2px solid transparent;margin-bottom:-1px;display:inline-flex;align-items:center;gap:8px}
    .request-tabs button:hover{color:#2f4d70}
    .request-tabs button.selected{color:#244c78;border-bottom-color:#244c78}
    .tab-count{background:#eef2f7;color:#5d7088;border-radius:10px;padding:2px 8px;font-size:12px}
    .request-tabs button.selected .tab-count{background:#244c78;color:#fff}
    .request-filters{padding-top:17px}
    .op-filter{width:auto;min-width:180px}
    .op-badge{font-size:12px;font-weight:800;padding:5px 9px;border-radius:6px;display:inline-block;letter-spacing:.3px}
    .op-alta{background:#e9f1fc;color:#3c6ead}
    .op-modificacion{background:#f0edfc;color:#7059b3}
    .op-baja{background:#fdeced;color:#bd5056}
    .person-cell{display:flex;align-items:center;gap:9px}
    .mini-avatar{width:28px;height:28px;border-radius:50%;background:#e4edf8;color:#2f5580;display:grid;place-items:center;font-size:12px;font-weight:800;flex-shrink:0}
    .clickable-row{cursor:pointer}
    .action-cell{text-align:right}
    .btn-small{padding:7px 12px;font-size:13px}
    .resolution-cell{max-width:260px;overflow:hidden;text-overflow:ellipsis}
    .muted{color:#9aa6b6;font-weight:400}
    .inline-error{margin:0 22px 14px;background:#fdeced;border:1px solid #f5cfd2;color:#b2444b;border-radius:8px;padding:10px 13px;font-size:14px}
    .form-section .inline-error{margin:12px 0 0}
    .modal-title-row{display:flex;align-items:center;gap:10px;margin-bottom:6px}
    .modal-title-row h2{margin:0!important}
    .request-summary{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:16px;padding:20px 24px;background:#f9fbfd;border-bottom:1px solid #edf0f4}
    .request-summary div{display:flex;flex-direction:column;gap:6px;align-items:flex-start;min-width:0}
    .request-summary small{font-size:12px;color:#8996a7;font-weight:700;letter-spacing:.3px}
    .request-summary strong{font-size:14px;color:#2f4359;overflow-wrap:anywhere}
    .form-section+.form-section{padding-top:0}
    .section-hint{font-size:13px;color:#8491a3;margin:-8px 0 12px}
    .compare-table{border:1px solid #e8edf3;border-radius:8px}
    .compare-table table{white-space:normal}
    .compare-table td{vertical-align:top}
    .field-label{font-weight:700;color:#5b6c82!important;width:200px}
    .old-value{color:#97a3b3!important}
    .changed-row{background:#fff9ec}
    .changed-row:hover{background:#fff4dc!important}
    .changed-row .old-value{text-decoration:line-through}
    .changed-row .new-value{color:#8a5a12!important;font-weight:700}
    .note-box{background:#f6f8fb;border:1px solid #e6ebf2;border-radius:8px;padding:12px 14px;font-size:14px;color:#40536b;white-space:pre-line}
    .warning-box,.info-box,.resolution-box{border-radius:10px;padding:14px 16px;font-size:14px}
    .warning-box{background:#fdf1f1;border:1px solid #f3d3d5;color:#9c3f45}
    .info-box{background:#f1f6fc;border:1px solid #d8e5f4;color:#36577e}
    .warning-box p,.info-box p,.resolution-box p{margin:6px 0 0;font-size:13px;opacity:.9}
    .resolution-box.approved{background:#eef8f2;border:1px solid #cdebd9;color:#1f7a53}
    .resolution-box.rejected{background:#fdf1f1;border:1px solid #f3d3d5;color:#a6434a}
    .resolution-reason{margin-top:12px;background:#ffffffb3;border-radius:7px;padding:10px 12px;color:#3d4e63;white-space:pre-line}
    .resolution-reason small{display:block;font-size:12px;font-weight:700;color:#8291a4;margin-bottom:4px}
    .decision-section{border-top:1px solid #edf0f4;padding-top:20px!important}
    .decision-section label{display:block;font-size:12px;font-weight:700;color:#64748a;margin-bottom:7px}
    .decision-section textarea{resize:vertical}
    .input-error{border-color:#d1666c!important;box-shadow:0 0 0 3px #d1666c1c!important}
    .field-footer{display:flex;justify-content:space-between;gap:10px;margin-top:6px}
    .field-error{font-size:13px;color:#bd4b52}
    .char-count{font-size:12px;color:#9aa6b6;margin-left:auto}
    .btn-success,.btn-danger{border-radius:8px;padding:11px 16px;font-size:14px;font-weight:700;cursor:pointer;display:inline-flex;align-items:center;gap:8px;transition:.2s}
    .btn-success{background:#21835a;color:#fff;border:1px solid #21835a}
    .btn-success:hover{background:#1a6d4a}
    .btn-danger{background:#fff;color:#bd4b52;border:1px solid #efc4c7}
    .btn-danger:hover{background:#fdf1f1;border-color:#e3a3a8}
    .btn-success:disabled,.btn-danger:disabled,.btn-outline:disabled{opacity:.55;cursor:not-allowed}
    @media(max-width:900px){.request-summary{grid-template-columns:repeat(2,minmax(0,1fr))}}
    @media(max-width:700px){.request-tabs{padding:10px 12px 0}.request-tabs button{padding:10px 9px}.request-filters{flex-wrap:wrap}.op-filter{flex:1}.request-summary{padding:16px 17px}.heading-actions{flex-wrap:wrap}.segmented{width:100%}.segmented button{flex:1}.field-label{width:auto}.modal-footer{flex-wrap:wrap}.modal-footer button{flex:1}}
    @media(max-width:1100px){.stats-grid{grid-template-columns:repeat(2,minmax(0,1fr))}.content-area{padding:24px 20px}}
    @media(max-width:700px){.topbar{padding:0 15px}.content-area{padding:20px 12px}.page-heading{align-items:flex-start;flex-direction:column}.heading-actions{width:100%}.heading-actions button{flex:1;padding:10px 8px}.stats-grid{gap:9px}.stat-card{padding:13px}.stat-card>strong{font-size:25px}.inventory-toolbar{padding:17px 14px;align-items:flex-start;flex-direction:column}.search-row{padding:0 14px 14px}.table-footer{padding:13px}.form-grid{grid-template-columns:1fr}.detail-list{grid-template-columns:1fr}.modal-backdrop{padding:9px}.modal-card{max-height:95vh}.modal-header,.form-section{padding:17px}.modal-footer{padding:13px 17px}.upload-zone{margin:10px 17px 18px}.preview-section{padding:0 17px 17px}}
  `]
})
export class AppComponent {

  api = API_URL;

  logged = !!localStorage.getItem('dozzer_token');
  email = localStorage.getItem('dozzer_email') || '';
  roles: string[] = JSON.parse(localStorage.getItem('dozzer_roles') || '[]');
  permissions: string[] = JSON.parse(localStorage.getItem('dozzer_permissions') || '[]');
  isAdmin = this.roles.includes('ADMIN');
  canAuthorize = this.permissions.includes('SOLICITUDES_AUTORIZAR');
  nombreUsuario = localStorage.getItem('dozzer_nombre') || '';

  // Cambio de contraseña (obligatorio con contraseña temporal, o voluntario)
  mustChangePassword = localStorage.getItem('dozzer_cambio_password') === '1';
  cambioPasswordModal = false;
  pwd = { actual: '', nueva: '', confirmar: '' };
  pwdError = '';
  pwdGuardando = false;

  // Bandejas de solicitudes
  solicitudes: any[] = [];
  vistaSolicitudes: 'autorizacion' | 'mias' = (this.isAdmin || this.canAuthorize) ? 'autorizacion' : 'mias';
  tabSolicitudes = (this.isAdmin || this.canAuthorize) ? 'PENDIENTE' : 'TODAS';
  filtroSolicitudes = '';
  filtroOperacion = '';
  solicitudesLoading = false;
  solicitudesError = '';
  solicitudSeleccionada: any = null;
  camposSolicitud: { etiqueta: string; actual: string; propuesto: string; cambio: boolean }[] = [];
  hayValoresActuales = false;
  motivoResolucion = '';
  motivoRequerido = false;
  errorResolucion = '';
  resolviendo: '' | 'aprobar' | 'rechazar' = '';
  modoRevocar = false;
  motivoRevocacion = '';
  revocacionRequerida = false;
  errorRevocacion = '';
  revocando = false;

  tabsSolicitudes = [
    { key: 'PENDIENTE', label: 'Pendientes' },
    { key: 'APROBADA', label: 'Aprobadas' },
    { key: 'RECHAZADA', label: 'Rechazadas' },
    { key: 'TODAS', label: 'Todas' }
  ];

  // Campos del bien: clave enviada en la solicitud / clave en el inventario actual
  camposBien: any[] = [
    { clave: 'clave', claveActual: 'clave', etiqueta: 'Clave', siempre: true },
    { clave: 'articulo', claveActual: 'articulo', etiqueta: 'Nombre del bien', siempre: true },
    { clave: 'idClasificacion', claveActual: 'id_clasificacion', etiqueta: 'Clasificación', catalogo: 'clasificaciones', idCatalogo: 'id_clasificacion', nombreCatalogo: 'nombre' },
    { clave: 'idSubcategoria', claveActual: 'id_subcategoria', etiqueta: 'Subcategoría', catalogo: 'subcategorias', idCatalogo: 'id_subcategoria', nombreCatalogo: 'nombre_subcategoria' },
    { clave: 'idMarca', claveActual: 'id_marca', etiqueta: 'Marca', catalogo: 'marcas', idCatalogo: 'id_marca', nombreCatalogo: 'nombre_marca' },
    { clave: 'modelo', claveActual: 'modelo', etiqueta: 'Modelo' },
    { clave: 'numeroSerie', claveActual: 'numero_serie', etiqueta: 'Número de serie' },
    { clave: 'idColor', claveActual: 'id_color', etiqueta: 'Color', catalogo: 'colores', idCatalogo: 'id_color', nombreCatalogo: 'nombre_color' },
    { clave: 'unidadMedida', claveActual: 'unidad_medida', etiqueta: 'Unidad de medida' },
    { clave: 'cantidadPiezas', claveActual: 'cantidad_piezas', etiqueta: 'Cantidad de piezas' },
    { clave: 'idResponsableArea', claveActual: 'id_responsable_area', etiqueta: 'Responsable de área', catalogo: 'personas', idCatalogo: 'id_persona', nombreCatalogo: 'nombre_completo' },
    { clave: 'idUsuarioDirecto', claveActual: 'id_usuario_directo', etiqueta: 'Usuario directo', catalogo: 'personas', idCatalogo: 'id_persona', nombreCatalogo: 'nombre_completo' },
    { clave: 'idUbicacion', claveActual: 'id_ubicacion', etiqueta: 'Ubicación', catalogo: 'ubicaciones', idCatalogo: 'id_ubicacion', nombreCatalogo: 'nombre_area' },
    { clave: 'estadoFisico', claveActual: 'estado_fisico', etiqueta: 'Estado físico' },
    { clave: 'estatusInventario', claveActual: 'estatus_inventario', etiqueta: 'Estatus de inventario' },
    { clave: 'estatusAsignacion', claveActual: 'estatus_asignacion', etiqueta: 'Estatus de asignación' },
    { clave: 'caracteristicas', claveActual: 'caracteristicas', etiqueta: 'Características' },
    { clave: 'requiereCalibracion', claveActual: 'requiere_calibracion', etiqueta: 'Requiere calibración' },
    { clave: 'fechaCalibracion', claveActual: 'fecha_calibracion', etiqueta: 'Próxima calibración' },
    { clave: 'requiereMantenimiento', claveActual: 'requiere_mantenimiento', etiqueta: 'Requiere mantenimiento' },
    { clave: 'fechaMantenimiento', claveActual: 'fecha_mantenimiento', etiqueta: 'Próximo mantenimiento' }
  ];

  section = 'inicio';
  menuCollapsed = true;

  error = '';
  aviso: { tipo: 'exito' | 'pendiente' | 'baja' | 'rechazo' | 'error' | 'info'; titulo: string; mensaje: string; ms: number;
           accion?: { texto: string; fn: () => void } } | null = null;

  // Formatos PDF de Alta / Baja
  documento: { datos: FormatoDoc; url: string; vista: SafeResourceUrl; blob: Blob; nombre: string } | null = null;
  docGenerando = false;
  waNumero = '';
  vistaDoc: 'inicio' | 'whatsapp' = 'inicio';
  waUrl = '';

  // Inventario anual
  levantamientos: any[] = [];
  levCargando = false;
  levAbierto: any = null;
  levDetalle: any = null;
  anualVista: 'lista' | 'captura' | 'avance' = 'lista';
  // Captura: "bien nuevo", "buscar en todas las sedes" y opciones de la sede (áreas, personas, catálogos)
  capNuevo: any = null;
  capBuscar: any = null;
  capOpciones: any = { areas: [], personas: [], clasificaciones: [], marcas: [] };
  av: any = null;
  private avTimer: any = null;
  readonly avFiltros = [
    { key: 'todos', label: 'Todos', ico: '' },
    { key: 'pendientes', label: 'Pendientes', ico: '⏳' },
    { key: 'noenc', label: 'No encontrados', ico: '✗' },
    { key: 'revisados', label: 'Revisados', ico: '✓' },
    { key: 'fuera', label: 'Fuera de lugar', ico: '⚠' },
    { key: 'sinfoto', label: 'Sin foto', ico: '📷' },
    { key: 'sinetiq', label: 'Sin etiqueta', ico: '🏷' }
  ];
  levNuevo: any = null;
  levGuardando = false;
  levError = '';
  candidatos: any[] = [];
  cap: any = null;
  @ViewChild('camQr') camQr?: ElementRef<HTMLVideoElement>;
  private camStream: MediaStream | null = null;
  private camTimer: any = null;

  // Pantalla de Documentos
  docsVista: 'menu' | 'ALTA' | 'BAJA' | 'ACTAS' | 'RESGUARDO' = 'menu';
  documentosLista: any[] = [];
  docsCargando = false;
  docsError = '';
  docsBusqueda = '';
  abriendoDocumento: any = null;
  waAviso = '';
  docMensaje = '';
  docError = '';
  private logoDataUrl: string | null = null;
  private avisoTimer: any;
  filter = '';
  statusFilter = '';

  loading = false;
  saving = false;

  editing = false;
  // Inventario
  bienesCargando = false;
  bienesError = '';
  bienesVista: any[] = [];
  bienSeleccionado: any = null;
  fotoAmpliada: { url: string; titulo: string } | null = null;
  filtroSede: number | null = null;
  filtroClasificacion: number | null = null;
  filtroEstadoFisico = '';
  filtroAsignacion = '';
  filtroCM: '' | 'atencion' | 'calibracion' | 'mantenimiento' | 'sinfecha' = '';
  ordenBienes: { key: string; dir: 'asc' | 'desc' } = { key: '', dir: 'asc' };
  paginaActualBienes = 1;
  tamanoPaginaBienes = 25;
  totalPiezas = 0;
  bienesBuenEstado = 0;
  bienesDisponibles = 0;
  resumenClasificaciones: { id: any; nombre: string; total: number; pct: number }[] = [];
  columnPickerOpen = false;
  private busquedaTimer: any;

  // Diálogo de confirmación con diseño propio (reemplaza al confirm() del navegador)
  dialogo: DialogoConfirmacion | null = null;
  bienOriginal: any = null;

  // Etiquetas QR
  qrPreview = '';
  qrMuestra = '';
  modalEtiquetas = false;
  etiquetasAlcance: 'filtrados' | 'pagina' = 'filtrados';
  etiquetasFormato: 'estandar' | 'compacta' = 'estandar';
  generandoEtiquetas = false;

  // Columnas que se pueden mostrar en la tabla de inventario
  columnasDisponibles: { key: string; label: string; tipo?: string; fija?: boolean }[] = [
    { key: 'clave', label: 'Clave', tipo: 'clave', fija: true },
    { key: 'articulo', label: 'Bien', tipo: 'bien', fija: true },
    { key: 'marca', label: 'Marca' },
    { key: 'modelo', label: 'Modelo' },
    { key: 'numero_serie', label: 'No. de serie' },
    { key: 'clasificacion', label: 'Clasificación' },
    { key: 'subcategoria', label: 'Subcategoría' },
    { key: 'color', label: 'Color' },
    { key: 'sede', label: 'Sede' },
    { key: 'ubicacion', label: 'Ubicación' },
    { key: 'responsable', label: 'Responsable', tipo: 'persona' },
    { key: 'usuario_directo', label: 'Usuario directo', tipo: 'persona' },
    { key: 'cantidad_piezas', label: 'Piezas', tipo: 'num' },
    { key: 'estado_fisico', label: 'Estado físico', tipo: 'fisico' },
    { key: 'estatus_asignacion', label: 'Asignación' },
    { key: 'estatus_inventario', label: 'Estatus', tipo: 'estatus' },
    { key: 'fecha_alta', label: 'Fecha de alta', tipo: 'fecha' }
  ];
  columnasElegidas: string[] = this.cargarColumnas();

  excelModal = false;
  excelFile: File | null = null;
  excelRows: any[] = [];
  excelHeaders: string[] = [];
  excelError = '';
  excelImporting = false;
  excelImportResult = '';

  verPasswordLogin = false;
  // Recuperación de contraseña
  modoRecuperar = false;
  correoRecuperar = '';
  numeroRecuperar = '';
  recuperacionEnviando = false;
  nuevaRecuperar = '';
  confirmarRecuperar = '';
  recuperacionError = '';
  restablecerManualAbierto = false;
  correoRestablecer = '';
  passwordTemporal: any = null;
  temporalCopiada = false;
  menuMasInventario = false;
  menuUsuario = false;
  // App instalable y avisos al celular (Web Push)
  pushEstado: 'cargando' | 'no-soportado' | 'sin-https' | 'instalar-ios' | 'bloqueado' | 'inactivo' | 'activo' = 'cargando';
  pushProcesando = false;
  instalarPrompt: any = null;
  avisoPushOculto = this.leerLocal('dozzer_push_oculto') === '1';
  private swRegistro: Promise<ServiceWorkerRegistration | null> = Promise.resolve(null);
  vistaInv: 'menu' | 'sedes' | 'bienes' | 'personas' | 'completo' = 'menu';
  origenVistaInv: 'menu' | 'sedes' | 'personas' = 'menu';
  filtroPersona: any = null;
  busquedaPersonas = '';
  personasVisibles = 48;
  tarjetasVisibles = 36;
  resumenSedes: any[] = [];
  resumenPersonas: any[] = [];
  coloresClasificacion: Record<string, string> = {};
  readonly anioActual = new Date().getFullYear();

  login = {
    correo: '',
    password: ''
  };

  bienes: any[] = [];

  // Catálogos
  catalog = 'sedes';
  catForm: any = {};
  catalogFilter = '';
  catalogShowInactive = false;
  catalogPage = 1;
  catalogPageSize = 25;
  catalogError = '';
  catalogModal = false;
  catalogEditId: number | null = null;
  catalogSaving = false;
  catalogFormError = '';
  catalogFieldError = '';

  bien: any = {
    clave: '',
    articulo: '',
    cantidadPiezas: 1,
    estatusInventario: 'ACTIVO',
    estatusAsignacion: 'SIN_ASIGNAR'
  };

  catalogs: any = {
    sedes: [],
    ubicaciones: [],
    personas: [],
    clasificaciones: [],
    marcas: [],
    colores: [],
    departamentos: [],
    puestos: [],
    entidades: [],
    'tipos-ubicacion': [],
    subcategorias: []
  };

  // Definición de cada catálogo: columnas de la tabla y campos del formulario.
  // En los select, "opciones" indica el catálogo de origen y los campos de valor/texto.
  catalogConfigs: any = {
    entidades: {
      label: 'Entidades', singular: 'Entidad', id: 'id_entidad', estatus: true,
      descripcion: 'Estados de la república donde la empresa tiene sedes.',
      columnas: [{ key: 'nombre_entidad', label: 'Entidad' }, { key: 'abreviatura', label: 'Abreviatura' }],
      campos: [
        { key: 'nombre_entidad', label: 'Nombre de la entidad', requerido: true, max: 100, placeholder: 'Ej. GUANAJUATO' },
        { key: 'abreviatura', label: 'Abreviatura', requerido: true, max: 20, placeholder: 'Ej. GTO' },
        { key: 'estatus', label: 'Activo', tipo: 'check' }
      ]
    },
    sedes: {
      label: 'Sedes', singular: 'Sede', id: 'id_sede', estatus: true,
      descripcion: 'Oficinas o centros de trabajo. Cada bien, ubicación y persona pertenece a una sede.',
      columnas: [{ key: 'nombre_sede', label: 'Sede' }, { key: 'prefijo_clave', label: 'Prefijo de clave' }, { key: 'nombre_entidad', label: 'Entidad' }],
      campos: [
        { key: 'nombre_sede', label: 'Nombre de la sede', requerido: true, max: 150 },
        { key: 'prefijo_clave', label: 'Prefijo de clave', max: 20, placeholder: 'Ej. CE' },
        { key: 'direccion_completa', label: 'Dirección completa', max: 300, placeholder: 'Calle, número, colonia y código postal' },
        { key: 'entidad', label: 'Entidad', tipo: 'select', texto: true, opciones: { catalogo: 'entidades', valor: 'id_entidad', texto: 'nombre_entidad' } },
        { key: 'estatus', label: 'Activa', tipo: 'check' }
      ]
    },
    'tipos-ubicacion': {
      label: 'Tipos de ubicación', singular: 'Tipo de ubicación', id: 'id_tipo_ubicacion', estatus: true,
      descripcion: 'Clasificación de las áreas: almacén, oficina, campamento, etc.',
      columnas: [{ key: 'nombre_tipo_ubicacion', label: 'Tipo' }, { key: 'es_temporal', label: 'Temporal', bool: true }],
      campos: [
        { key: 'nombre_tipo_ubicacion', label: 'Nombre del tipo', requerido: true, max: 150, ancho: 'completo' },
        { key: 'es_temporal', label: 'Es una ubicación temporal (ej. campamento de obra)', tipo: 'check' },
        { key: 'estatus', label: 'Activo', tipo: 'check' }
      ]
    },
    ubicaciones: {
      label: 'Ubicaciones', singular: 'Ubicación', id: 'id_ubicacion', estatus: true,
      descripcion: 'Áreas físicas dentro de cada sede donde se encuentran los bienes.',
      columnas: [{ key: 'nombre_area', label: 'Área' }, { key: 'nombre_tipo_ubicacion', label: 'Tipo' }, { key: 'nombre_sede', label: 'Sede' }],
      campos: [
        { key: 'nombre_area', label: 'Nombre del área', requerido: true, max: 150 },
        { key: 'id_sede', label: 'Sede', tipo: 'select', requerido: true, opciones: { catalogo: 'sedes', valor: 'id_sede', texto: 'nombre_sede' } },
        { key: 'tipo_ubicacion', label: 'Tipo de ubicación', tipo: 'select', texto: true, opciones: { catalogo: 'tipos-ubicacion', valor: 'id_tipo_ubicacion', texto: 'nombre_tipo_ubicacion' } },
        { key: 'estatus', label: 'Activa', tipo: 'check' }
      ]
    },
    departamentos: {
      label: 'Departamentos', singular: 'Departamento', id: 'id_departamento', estatus: true,
      descripcion: 'Áreas organizacionales de la empresa. Agrupan los puestos.',
      columnas: [{ key: 'nombre_departamento', label: 'Departamento' }],
      campos: [
        { key: 'nombre_departamento', label: 'Nombre del departamento', requerido: true, max: 180, ancho: 'completo' },
        { key: 'estatus', label: 'Activo', tipo: 'check' }
      ]
    },
    puestos: {
      label: 'Puestos', singular: 'Puesto', id: 'id_puesto', estatus: true,
      descripcion: 'Puestos de trabajo y el departamento al que pertenecen.',
      columnas: [{ key: 'nombre_puesto', label: 'Puesto' }, { key: 'nombre_departamento', label: 'Departamento' }],
      campos: [
        { key: 'nombre_puesto', label: 'Nombre del puesto', requerido: true, max: 180 },
        { key: 'id_departamento', label: 'Departamento', tipo: 'select', requerido: true, opciones: { catalogo: 'departamentos', valor: 'id_departamento', texto: 'nombre_departamento' } },
        { key: 'estatus', label: 'Activo', tipo: 'check' }
      ]
    },
    personas: {
      label: 'Personas', singular: 'Persona', id: 'id_persona', estatus: true,
      descripcion: 'Empleados que pueden ser responsables o usuarios de un bien.',
      columnas: [
        { key: 'nombre_completo', label: 'Nombre' }, { key: 'numero_empleado', label: 'No. empleado' },
        { key: 'correo', label: 'Correo' }, { key: 'nombre_puesto', label: 'Puesto' }, { key: 'nombre_sede', label: 'Sede' },
        { key: 'es_encargado', label: 'Encargado', bool: true }
      ],
      campos: [
        { key: 'nombre_completo', label: 'Nombre completo', requerido: true, max: 180, ancho: 'completo' },
        { key: 'numero_empleado', label: 'Número de empleado', max: 20 },
        { key: 'correo', label: 'Correo electrónico', tipo: 'email', max: 180, placeholder: 'nombre@dozzer.mx' },
        { key: 'id_puesto', label: 'Puesto', tipo: 'select', opciones: { catalogo: 'puestos', valor: 'id_puesto', texto: 'nombre_puesto' } },
        { key: 'id_sede', label: 'Sede', tipo: 'select', requerido: true, opciones: { catalogo: 'sedes', valor: 'id_sede', texto: 'nombre_sede' } },
        { key: 'es_encargado', label: 'Encargado de inventario de su área', tipo: 'check' },
        { key: 'estatus', label: 'Activa', tipo: 'check' }
      ]
    },
    clasificaciones: {
      label: 'Clasificaciones', singular: 'Clasificación', id: 'id_clasificacion', estatus: false,
      descripcion: 'Categorías principales de los bienes.',
      columnas: [{ key: 'nombre', label: 'Clasificación' }, { key: 'codigo', label: 'Código' }],
      campos: [
        { key: 'nombre', label: 'Nombre', requerido: true, max: 150 },
        { key: 'codigo', label: 'Código', max: 40 }
      ]
    },
    subcategorias: {
      label: 'Subcategorías', singular: 'Subcategoría', id: 'id_subcategoria', estatus: false,
      descripcion: 'Detalle de cada clasificación (ej. Mobiliario, Herramienta).',
      columnas: [{ key: 'nombre_subcategoria', label: 'Subcategoría' }, { key: 'nombre_clasificacion', label: 'Clasificación' }],
      campos: [
        { key: 'nombre_subcategoria', label: 'Nombre de la subcategoría', requerido: true, max: 150 },
        { key: 'id_clasificacion', label: 'Clasificación', tipo: 'select', requerido: true, opciones: { catalogo: 'clasificaciones', valor: 'id_clasificacion', texto: 'nombre' } }
      ]
    },
    marcas: {
      label: 'Marcas', singular: 'Marca', id: 'id_marca', estatus: false,
      descripcion: 'Fabricantes de los bienes.',
      columnas: [{ key: 'nombre_marca', label: 'Marca' }],
      campos: [{ key: 'nombre_marca', label: 'Nombre de la marca', requerido: true, ancho: 'completo' }]
    },
    colores: {
      label: 'Colores', singular: 'Color', id: 'id_color', estatus: false,
      descripcion: 'Colores disponibles para describir los bienes.',
      columnas: [{ key: 'nombre_color', label: 'Color' }],
      campos: [{ key: 'nombre_color', label: 'Nombre del color', requerido: true, ancho: 'completo' }]
    }
  };

  catalogGroups = [
    { label: 'ORGANIZACIÓN', items: [
      { key: 'entidades', label: 'Entidades' }, { key: 'sedes', label: 'Sedes' },
      { key: 'tipos-ubicacion', label: 'Tipos de ubicación' }, { key: 'ubicaciones', label: 'Ubicaciones' }
    ] },
    { label: 'PERSONAL', items: [
      { key: 'departamentos', label: 'Departamentos' }, { key: 'puestos', label: 'Puestos' }, { key: 'personas', label: 'Personas' }
    ] },
    { label: 'BIENES', items: [
      { key: 'clasificaciones', label: 'Clasificaciones' }, { key: 'subcategorias', label: 'Subcategorías' },
      { key: 'marcas', label: 'Marcas' }, { key: 'colores', label: 'Colores' }
    ] }
  ];

  constructor(private http: HttpClient, private sanitizer: DomSanitizer) {
    this.prepararApp();
    if (this.logged && !this.mustChangePassword) {
      this.loadAll();
    }
  }

  // =========================
  // APP INSTALABLE Y AVISOS AL CELULAR
  // =========================

  private leerLocal(k: string): string | null {
    try { return localStorage.getItem(k); } catch { return null; }
  }

  private prepararApp(): void {
    // Un aviso abrió la app con ?ir=autorizaciones|solicitudes
    try {
      const params = new URLSearchParams(window.location.search);
      const ir = params.get('ir');
      if (ir) {
        history.replaceState(null, '', window.location.pathname);
        setTimeout(() => this.irDesdeAviso(ir));
      }
      const etiqueta = this.codigoDeTexto(params.get('e') || '');
      if (etiqueta) {
        history.replaceState(null, '', window.location.pathname);
        try { sessionStorage.setItem('dozzer_etiqueta', etiqueta); } catch { /* */ }
      }
    } catch { /* sin parámetros */ }
    window.addEventListener('beforeinstallprompt', (e: any) => { e.preventDefault(); this.instalarPrompt = e; });
    window.addEventListener('appinstalled', () => this.instalarPrompt = null);
    if ('serviceWorker' in navigator && window.isSecureContext) {
      this.swRegistro = navigator.serviceWorker.register('sw-dozzer.js').catch(() => null);
      // Con la app abierta, tocar un aviso la lleva a la pantalla que corresponde
      navigator.serviceWorker.addEventListener('message', (ev: any) => {
        if (ev.data?.tipo !== 'dozzer-ir') return;
        try { const ir = new URL(ev.data.url).searchParams.get('ir'); if (ir) this.irDesdeAviso(ir); } catch { /* */ }
      });
    }
  }

  private irDesdeAviso(ir: string): void {
    if (!this.logged || this.mustChangePassword) return;
    if (ir === 'autorizaciones' || ir === 'solicitudes') {
      this.section = 'solicitudes';
      this.menuCollapsed = true;
      this.loadSolicitudes();
    }
  }

  get textoPushEstado(): string {
    switch (this.pushEstado) {
      case 'activo': return 'En este dispositivo · toca para apagarlos';
      case 'inactivo': return this.isAdmin || this.canAuthorize ? 'Que suene con autorizaciones pendientes' : 'Cuando respondan tus solicitudes';
      case 'instalar-ios': return 'Primero agrégala a la pantalla de inicio';
      case 'bloqueado': return 'Bloqueados en este navegador';
      case 'sin-https': return 'Solo en la versión de internet';
      case 'no-soportado': return 'Este navegador no los permite';
      default: return 'Revisando…';
    }
  }

  // Al entrar: ¿este dispositivo ya tiene avisos? Si sí, se liga al usuario que inició sesión
  async revisarPush(): Promise<void> {
    const ios = /iPhone|iPad|iPod/i.test(navigator.userAgent);
    const instalada = window.matchMedia?.('(display-mode: standalone)').matches || (navigator as any).standalone === true;
    if (!window.isSecureContext) { this.pushEstado = 'sin-https'; return; }
    if (!('serviceWorker' in navigator) || !('PushManager' in window) || !('Notification' in window)) {
      this.pushEstado = ios && !instalada ? 'instalar-ios' : 'no-soportado';
      return;
    }
    if (Notification.permission === 'denied') { this.pushEstado = 'bloqueado'; return; }
    try {
      const reg = await this.swRegistro;
      if (!reg) { this.pushEstado = 'no-soportado'; return; }
      const sub = await reg.pushManager.getSubscription();
      if (sub && Notification.permission === 'granted') {
        await firstValueFrom(this.http.post(this.api + '/push/suscripciones', sub.toJSON()));
        this.pushEstado = 'activo';
      } else {
        this.pushEstado = 'inactivo';
      }
    } catch {
      this.pushEstado = 'inactivo';
    }
  }

  async activarPush(): Promise<void> {
    if (this.pushProcesando) return;
    const avisos: any = {
      'instalar-ios': ['Primero instala DOZZER', 'En el iPhone toca Compartir ⬆ y luego "Agregar a pantalla de inicio". Abre DOZZER desde ese ícono y activa los avisos ahí.'],
      'sin-https': ['Solo en la versión de internet', 'Los avisos funcionan al abrir DOZZER desde su dirección de internet (https), no desde la red de la oficina.'],
      'no-soportado': ['Este navegador no permite avisos', 'Usa Chrome en Android, o en iPhone (iOS 16.4 o más reciente) instala DOZZER en la pantalla de inicio.'],
      'bloqueado': ['Avisos bloqueados', 'Este navegador tiene bloqueadas las notificaciones de DOZZER. Permítelas en los ajustes del sitio (el candado junto a la dirección) y vuelve a intentar.']
    };
    if (avisos[this.pushEstado]) { this.notificar('info', avisos[this.pushEstado][0], avisos[this.pushEstado][1], 0); return; }
    this.pushProcesando = true;
    try {
      const permiso = await Notification.requestPermission();
      if (permiso !== 'granted') {
        this.pushEstado = permiso === 'denied' ? 'bloqueado' : 'inactivo';
        this.notificar('info', 'Avisos no activados', 'No se dio permiso para mostrar notificaciones.');
        return;
      }
      const reg = await this.swRegistro;
      if (!reg) throw new Error('sin service worker');
      const r: any = await firstValueFrom(this.http.get<any>(this.api + '/push/clave-publica'));
      const clave = this.b64aBytes(r.publicKey);
      let sub = await reg.pushManager.getSubscription();
      // Si el dispositivo estaba suscrito con otras llaves (otro servidor), se rehace
      if (sub && !this.mismaClave(sub, clave)) { await sub.unsubscribe(); sub = null; }
      if (!sub) sub = await reg.pushManager.subscribe({ userVisibleOnly: true, applicationServerKey: clave as BufferSource });
      await firstValueFrom(this.http.post(this.api + '/push/suscripciones', sub.toJSON()));
      this.pushEstado = 'activo';
      const prueba: any = await firstValueFrom(this.http.post<any>(this.api + '/push/prueba', {}));
      this.notificar('exito', 'Avisos activados', prueba?.enviados
        ? 'Te mandamos un aviso de prueba a este dispositivo.'
        : 'Quedaron activados. Si no llega el aviso de prueba, revisa que el celular permita notificaciones de DOZZER.');
    } catch (e) {
      this.notificar('error', 'No se activaron los avisos', this.mensajeError(e, 'Inténtalo de nuevo en unos segundos.'), 0);
    } finally {
      this.pushProcesando = false;
    }
  }

  async desactivarPush(): Promise<void> {
    if (this.pushProcesando) return;
    this.pushProcesando = true;
    try {
      const reg = await this.swRegistro;
      const sub = await reg?.pushManager.getSubscription();
      if (sub) {
        await firstValueFrom(this.http.post(this.api + '/push/suscripciones/quitar', { endpoint: sub.endpoint })).catch(() => null);
        await sub.unsubscribe();
      }
      this.pushEstado = 'inactivo';
      this.notificar('info', 'Avisos apagados', 'Este dispositivo ya no recibirá avisos de DOZZER.');
    } catch {
      this.notificar('error', 'No se apagaron los avisos', 'Inténtalo de nuevo.', 0);
    } finally {
      this.pushProcesando = false;
    }
  }

  // Al cerrar sesión, este dispositivo deja de recibir los avisos de esa persona
  private quitarPushAlSalir(): void {
    const token = this.leerLocal('dozzer_token');
    if (!token || this.pushEstado !== 'activo') return;
    this.swRegistro.then(reg => reg?.pushManager.getSubscription()).then(sub => {
      if (!sub) return;
      this.http.post(this.api + '/push/suscripciones/quitar', { endpoint: sub.endpoint },
        { headers: new HttpHeaders({ Authorization: 'Bearer ' + token }) }).subscribe({ error: () => {} });
    }).catch(() => {});
  }

  ocultarAvisoPush(): void {
    this.avisoPushOculto = true;
    try { localStorage.setItem('dozzer_push_oculto', '1'); } catch { /* opcional */ }
  }

  async instalarApp(): Promise<void> {
    const p = this.instalarPrompt;
    if (!p) return;
    this.instalarPrompt = null;
    try { await p.prompt(); } catch { /* cancelado */ }
  }

  private b64aBytes(b64: string): Uint8Array {
    const s = (b64 + '='.repeat((4 - b64.length % 4) % 4)).replace(/-/g, '+').replace(/_/g, '/');
    const bin = atob(s);
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }

  private mismaClave(sub: PushSubscription, clave: Uint8Array): boolean {
    const actual = sub.options?.applicationServerKey;
    if (!actual) return true;
    const a = new Uint8Array(actual);
    return a.length === clave.length && a.every((x, i) => x === clave[i]);
  }

  // =========================
  // CONTRASEÑA
  // =========================

  get pwdReglas() {
    const nueva = this.pwd.nueva || '';
    return {
      longitud: nueva.length >= 8,
      letrasNumeros: /[A-Za-zÁÉÍÓÚÑáéíóúñ]/.test(nueva) && /\d/.test(nueva),
      distinta: !!nueva && nueva !== this.pwd.actual,
      coincide: !!nueva && nueva === this.pwd.confirmar
    };
  }

  get pwdValida(): boolean {
    const r = this.pwdReglas;
    return !!this.pwd.actual && r.longitud && r.letrasNumeros && r.distinta && r.coincide;
  }

  abrirCambioPassword(): void {
    this.pwd = { actual: '', nueva: '', confirmar: '' };
    this.pwdError = '';
    this.cambioPasswordModal = true;
  }

  guardarPassword(): void {
    if (this.pwdGuardando) return;
    if (!this.pwdValida) {
      this.pwdError = 'Revisa que la nueva contraseña cumpla todos los requisitos.';
      return;
    }

    this.pwdGuardando = true;
    this.pwdError = '';
    this.http.post<any>(this.api + '/auth/cambiar-password', {
      passwordActual: this.pwd.actual,
      passwordNueva: this.pwd.nueva
    }).subscribe({
      next: () => {
        this.pwdGuardando = false;
        this.pwd = { actual: '', nueva: '', confirmar: '' };
        this.cambioPasswordModal = false;
        this.notificar('exito', 'Contraseña actualizada', 'Tu nueva contraseña ya está activa.');
        if (this.mustChangePassword) {
          this.mustChangePassword = false;
          localStorage.removeItem('dozzer_cambio_password');
          this.loadAll();
        }
      },
      error: e => {
        this.pwdGuardando = false;
        this.pwdError = this.mensajeError(e, 'No fue posible actualizar la contraseña.');
      }
    });
  }

  // Saludo de bienvenida
  get primerNombre(): string {
    const nombre = (this.nombreUsuario || '').trim().split(/\s+/)[0];
    if (nombre) return nombre.toUpperCase();
    const usuario = (this.email || '').split('@')[0].split(/[._-]/)[0];
    return usuario ? usuario.toUpperCase() : (this.isAdmin ? 'ADMINISTRADOR' : '');
  }

  get saludoHora(): string {
    const h = new Date().getHours();
    return h < 12 ? 'Buenos días' : h < 19 ? 'Buenas tardes' : 'Buenas noches';
  }

  get fechaHoy(): string {
    const texto = new Date().toLocaleDateString('es-MX', { weekday: 'long', day: 'numeric', month: 'long' });
    return texto.charAt(0).toUpperCase() + texto.slice(1);
  }

  get initials(): string {
    return this.email.split('@')[0].split(/[._-]/).map(x => x.charAt(0)).join('').substring(0, 2).toUpperCase() || 'AD';
  }

  get sectionTitle(): string {
    return ({
      inicio: 'Inicio',
      anual: 'Levantamiento de inventario',
      bienes: 'Bienes',
      documentos: 'Documentos',
      solicitudes: 'Solicitudes',
      catalogos: 'Configuración'
    } as any)[this.section];
  }

  get sectionLabel(): string {
    return ({
      inicio: 'INICIO',
      anual: 'LEVANTAMIENTO',
      bienes: 'BIENES',
      documentos: 'DOCUMENTOS',
      solicitudes: 'SOLICITUDES',
      catalogos: 'CONFIGURACIÓN'
    } as any)[this.section];
  }

  get catalogConfig(): any {
    return this.catalogConfigs[this.catalog];
  }

  get filteredCatalogRows(): any[] {
    const config = this.catalogConfig;
    const texto = this.catalogFilter.toLowerCase().trim();
    return (this.catalogs[this.catalog] || []).filter((x: any) => {
      if (config.estatus && !this.catalogShowInactive && x.estatus === false) return false;
      if (!texto) return true;
      return config.columnas.map((c: any) => this.valorColumna(x, c)).join(' ').toLowerCase().includes(texto);
    });
  }

  get catalogPages(): number {
    return Math.max(1, Math.ceil(this.filteredCatalogRows.length / this.catalogPageSize));
  }

  get pagedCatalogRows(): any[] {
    const pagina = Math.min(this.catalogPage, this.catalogPages);
    const inicio = (pagina - 1) * this.catalogPageSize;
    return this.filteredCatalogRows.slice(inicio, inicio + this.catalogPageSize);
  }

  // Subcategorías que corresponden a la clasificación elegida en el formulario de bienes
  get subcategoriasBien(): any[] {
    const lista = this.catalogs.subcategorias || [];
    return this.bien?.idClasificacion
      ? lista.filter((s: any) => s.id_clasificacion === this.bien.idClasificacion)
      : lista;
  }

  get columnasVisibles(): any[] {
    return this.columnasDisponibles.filter(c => this.columnaVisible(c.key));
  }

  get hayFiltrosBienes(): boolean {
    return !!(this.filter.trim() || this.filtrosActivos.length);
  }

  // Filtros aplicados, para mostrarlos como etiquetas que se pueden quitar
  get filtrosActivos(): { clave: string; etiqueta: string; valor: string }[] {
    const lista: { clave: string; etiqueta: string; valor: string }[] = [];
    if (this.filtroSede) {
      const sede = (this.catalogs.sedes || []).find((x: any) => x.id_sede === this.filtroSede);
      lista.push({ clave: 'sede', etiqueta: 'Sede', valor: sede?.nombre_sede || '#' + this.filtroSede });
    }
    if (this.filtroClasificacion) {
      const c = (this.catalogs.clasificaciones || []).find((x: any) => x.id_clasificacion === this.filtroClasificacion);
      lista.push({ clave: 'clasificacion', etiqueta: 'Clasificación', valor: c?.nombre || '#' + this.filtroClasificacion });
    }
    if (this.statusFilter) lista.push({ clave: 'estatus', etiqueta: 'Estatus', valor: this.textoCampo(this.statusFilter) });
    if (this.filtroEstadoFisico) lista.push({ clave: 'fisico', etiqueta: 'Estado físico', valor: this.textoCampo(this.filtroEstadoFisico) });
    if (this.filtroPersona) {
      const p = this.resumenPersonas.find(x => x.id === this.filtroPersona);
      lista.push({ clave: 'persona', etiqueta: 'Persona', valor: p?.nombre || '#' + this.filtroPersona });
    }
    if (this.filtroAsignacion) lista.push({ clave: 'asignacion', etiqueta: 'Asignación', valor: this.filtroAsignacion === 'ASIGNADO' ? 'Asignados' : 'Sin asignar' });
    if (this.filtroCM) lista.push({ clave: 'cm', etiqueta: 'Calibración/mant.', valor: ({ atencion: 'Vencidos o próximos', calibracion: 'Requieren calibración', mantenimiento: 'Requieren mantenimiento', sinfecha: 'Sin fecha' } as any)[this.filtroCM] });
    return lista;
  }

  quitarFiltro(clave: string): void {
    if (clave === 'sede') this.filtroSede = null;
    if (clave === 'clasificacion') this.filtroClasificacion = null;
    if (clave === 'estatus') this.statusFilter = '';
    if (clave === 'fisico') this.filtroEstadoFisico = '';
    if (clave === 'asignacion') this.filtroAsignacion = '';
    if (clave === 'persona') this.filtroPersona = null;
    if (clave === 'cm') this.filtroCM = '';
    this.aplicarFiltrosBienes();
  }

  get totalPaginasBienes(): number {
    return Math.max(1, Math.ceil(this.bienesVista.length / this.tamanoPaginaBienes));
  }

  get paginaBienes(): any[] {
    const pagina = Math.min(this.paginaActualBienes, this.totalPaginasBienes);
    const inicio = (pagina - 1) * this.tamanoPaginaBienes;
    return this.bienesVista.slice(inicio, inicio + this.tamanoPaginaBienes);
  }

  get rangoBienes(): { desde: number; hasta: number } {
    if (!this.bienesVista.length) return { desde: 0, hasta: 0 };
    const pagina = Math.min(this.paginaActualBienes, this.totalPaginasBienes);
    const desde = (pagina - 1) * this.tamanoPaginaBienes + 1;
    return { desde, hasta: Math.min(desde + this.tamanoPaginaBienes - 1, this.bienesVista.length) };
  }

  doLogin(): void {
    this.error = '';
    this.loading = true;

    this.http.post<any>(this.api + '/auth/login', this.login).subscribe({
      next: r => {
        localStorage.setItem('dozzer_token', r.token);
        localStorage.setItem('dozzer_email', r.correo);
        this.roles = Array.isArray(r.roles) ? r.roles : [];
        localStorage.setItem('dozzer_roles', JSON.stringify(this.roles));
        this.permissions = Array.isArray(r.permisos) ? r.permisos : [];
        localStorage.setItem('dozzer_permissions', JSON.stringify(this.permissions));
        this.isAdmin = this.roles.includes('ADMIN');
        this.canAuthorize = this.permissions.includes('SOLICITUDES_AUTORIZAR');
        this.columnasElegidas = this.cargarColumnas();
        this.vistaSolicitudes = this.esRevisor ? 'autorizacion' : 'mias';
        this.tabSolicitudes = this.esRevisor ? 'PENDIENTE' : 'TODAS';
        this.nombreUsuario = r.nombre || '';
        localStorage.setItem('dozzer_nombre', this.nombreUsuario);
        this.mustChangePassword = r.debeCambiarPassword === true;
        if (this.mustChangePassword) {
          localStorage.setItem('dozzer_cambio_password', '1');
          // La contraseña temporal recién escrita se usa como "contraseña actual"
          this.pwd = { actual: this.login.password, nueva: '', confirmar: '' };
          this.pwdError = '';
        } else {
          localStorage.removeItem('dozzer_cambio_password');
        }
        this.logged = true;
        this.email = r.correo;
        this.loading = false;
        this.login.password = '';
        if (!this.mustChangePassword) this.loadAll();
      },
      error: e => {
        this.loading = false;
        this.error = e.error?.error || 'No fue posible iniciar sesión';
      }
    });
  }

  logout(): void {
    this.quitarPushAlSalir();
    this.pushEstado = 'cargando';
    localStorage.removeItem('dozzer_token');
    localStorage.removeItem('dozzer_email');
    localStorage.removeItem('dozzer_roles');
    localStorage.removeItem('dozzer_permissions');
    localStorage.removeItem('dozzer_nombre');
    localStorage.removeItem('dozzer_cambio_password');
    this.nombreUsuario = '';
    this.mustChangePassword = false;
    this.cambioPasswordModal = false;
    this.pwd = { actual: '', nueva: '', confirmar: '' };
    this.pwdError = '';
    this.roles = [];
    this.permissions = [];
    this.isAdmin = false;
    this.canAuthorize = false;
    this.logged = false;
    this.email = '';
    this.solicitudes = [];
    this.solicitudSeleccionada = null;
    this.bienes = [];
    this.bienesVista = [];
    this.bienSeleccionado = null;
    this.dialogo = null;
    this.filter = '';
    this.statusFilter = '';
    this.filtroSede = null;
    this.filtroClasificacion = null;
    this.filtroEstadoFisico = '';
    this.filtroAsignacion = '';
    this.filtroPersona = null;
    this.filtroCM = '';
    this.vistaInv = 'menu';
    this.filtroSolicitudes = '';
    this.filtroOperacion = '';
    this.section = 'inicio';
    this.menuCollapsed = true;
    this.detenerEscaner();
    this.levantamientos = [];
    this.cap = null;
    this.detenerTimerAvance();
    this.av = null;
    this.anualVista = 'lista';
  }

  loadAll(): void {
    this.revisarPush();
    let etiqueta: string | null = null;
    try { etiqueta = sessionStorage.getItem('dozzer_etiqueta'); sessionStorage.removeItem('dozzer_etiqueta'); } catch { /* */ }
    if (etiqueta) setTimeout(() => this.abrirEtiquetaDesdeUrl(etiqueta!), 600);
    this.loadLevantamientos();
    this.loadBienes();
    Object.keys(this.catalogs).forEach(k => this.loadCatalog(k));
    this.loadSolicitudes();
  }

  // =========================
  // RECUPERACIÓN DE CONTRASEÑA
  // =========================

  abrirRecuperacion(): void {
    this.modoRecuperar = true;
    this.correoRecuperar = this.login.correo || '';
    this.numeroRecuperar = '';
    this.nuevaRecuperar = '';
    this.confirmarRecuperar = '';
    this.recuperacionError = '';
    this.error = '';
  }

  cerrarRecuperacion(): void {
    this.modoRecuperar = false;
    this.nuevaRecuperar = '';
    this.confirmarRecuperar = '';
    if (this.correoRecuperar && !this.login.correo) this.login.correo = this.correoRecuperar;
  }

  tieneLetrasYNumeros(texto: string): boolean {
    return /[A-Za-zÁÉÍÓÚÑáéíóúñ]/.test(texto || '') && /\d/.test(texto || '');
  }

  // Restablece la contraseña con el correo + número de empleado y entra directamente
  solicitarRecuperacion(): void {
    const correo = (this.correoRecuperar || '').trim();
    const nueva = this.nuevaRecuperar || '';
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(correo)) { this.recuperacionError = 'Escribe un correo electrónico válido.'; return; }
    const numeroEmpleado = (this.numeroRecuperar || '').trim();
    if (!numeroEmpleado) { this.recuperacionError = 'Escribe tu número de empleado.'; return; }
    if (nueva.length < 8) { this.recuperacionError = 'La contraseña debe tener al menos 8 caracteres.'; return; }
    if (!this.tieneLetrasYNumeros(nueva)) { this.recuperacionError = 'La contraseña debe incluir letras y números.'; return; }
    if (nueva !== this.confirmarRecuperar) { this.recuperacionError = 'Las contraseñas no coinciden.'; return; }

    this.recuperacionEnviando = true;
    this.recuperacionError = '';
    this.http.post<any>(this.api + '/auth/recuperar-password', { correo, numeroEmpleado, passwordNueva: nueva }).subscribe({
      next: () => {
        this.recuperacionEnviando = false;
        this.modoRecuperar = false;
        this.login = { ...this.login, correo, password: nueva };
        this.nuevaRecuperar = '';
        this.confirmarRecuperar = '';
        this.numeroRecuperar = '';
        this.doLogin();
      },
      error: e => {
        this.recuperacionEnviando = false;
        this.recuperacionError = this.mensajeError(e, 'No fue posible restablecer la contraseña. Intenta de nuevo.');
      }
    });
  }

  // Asigna una contraseña temporal distinta a cada empleado que aún no cambia la suya y descarga la lista.
  generarTemporalesPendientes(): void {
    this.abrirDialogo({
      tipo: 'peligro', icono: '⚿',
      titulo: '¿Generar contraseñas temporales?',
      mensaje: 'Cada empleado que todavía no ha cambiado su contraseña recibirá una temporal distinta y se descargará un archivo de Excel con la lista. ' +
        'Las contraseñas anteriores dejarán de funcionar. Guarda el archivo en un lugar seguro y entrégale a cada persona solo la suya.',
      textoConfirmar: 'Sí, generar y descargar',
      accion: terminar => {
        this.http.post<any>(this.api + '/usuarios/recuperaciones/temporales-pendientes', {}).subscribe({
          next: r => {
            terminar();
            const filas = r?.usuarios || [];
            if (!filas.length) {
              this.notificar('exito', 'Nada que generar', 'Todos los empleados ya cambiaron su contraseña.');
              return;
            }
            const celda = (v: any) => '"' + String(v ?? '').replace(/"/g, '""') + '"';
            const csv = ['Sede,Nombre,Número de empleado,Correo,Contraseña temporal']
              .concat(filas.map((u: any) => [u.sedes, u.nombre, u.numeroEmpleado, u.correo, u.passwordTemporal].map(celda).join(',')))
              .join(String.fromCharCode(13, 10));
            const blob = new Blob([String.fromCharCode(0xFEFF) + csv], { type: 'text/csv;charset=utf-8' });
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = 'dozzer-contrasenas-temporales-' + new Date().toISOString().slice(0, 10) + '.csv';
            document.body.appendChild(a); a.click(); a.remove();
            setTimeout(() => URL.revokeObjectURL(a.href), 2000);
            this.notificar('exito', 'Contraseñas generadas', 'Se descargó la lista con ' + filas.length + ' empleados. Al entrar, cada uno deberá crear su propia contraseña.', 6000);
          },
          error: e => terminar(this.mensajeError(e, 'No fue posible generar las contraseñas.'))
        });
      }
    });
  }

  abrirRestablecerManual(): void {
    this.correoRestablecer = '';
    this.restablecerManualAbierto = true;
  }

  restablecerPassword(correo: string, nombre?: string): void {
    const destino = (correo || '').trim();
    if (!destino) return;
    this.abrirDialogo({
      tipo: 'confirmar',
      icono: '⚿',
      titulo: '¿Generar contraseña temporal?',
      mensaje: 'Se reemplazará la contraseña actual de ' + (nombre || destino) + '. Deberá crear una nueva al iniciar sesión.',
      textoConfirmar: 'Sí, generar',
      accion: terminar => {
        this.http.post<any>(this.api + '/usuarios/recuperaciones/restablecer', { correo: destino }).subscribe({
          next: r => {
            terminar();
            this.restablecerManualAbierto = false;
            this.temporalCopiada = false;
            this.passwordTemporal = r;
          },
          error: e => terminar(this.mensajeError(e, 'No fue posible restablecer la contraseña.'))
        });
      }
    });
  }

  copiarTemporal(): void {
    const texto = this.passwordTemporal?.passwordTemporal;
    if (!texto) return;
    const listo = () => { this.temporalCopiada = true; setTimeout(() => this.temporalCopiada = false, 2000); };
    try {
      navigator.clipboard.writeText(texto).then(listo).catch(() => this.copiarRespaldo(texto, listo));
    } catch {
      this.copiarRespaldo(texto, listo);
    }
  }

  // Copiado alterno para páginas abiertas por IP (sin https el portapapeles moderno no está disponible)
  private copiarRespaldo(texto: string, listo: () => void): void {
    const area = document.createElement('textarea');
    area.value = texto;
    area.style.position = 'fixed';
    area.style.opacity = '0';
    document.body.appendChild(area);
    area.select();
    try { if (document.execCommand('copy')) listo(); } catch { /* sin portapapeles */ }
    document.body.removeChild(area);
  }

  inicialesDe(texto: string): string {
    const partes = String(texto || '').replace(/@.*/, '').split(/[\s._-]+/).filter(Boolean);
    return ((partes[0]?.[0] || '') + (partes[1]?.[0] || '')).toUpperCase() || '?';
  }

  // =========================
  // BANDEJAS DE SOLICITUDES
  // =========================

  get esRevisor(): boolean {
    return this.isAdmin || this.canAuthorize;
  }

  get pendientesBadge(): number {
    const solicitudes = this.vistaSolicitudes === 'autorizacion' ? this.conteoSolicitudes('PENDIENTE') : 0;
    return solicitudes;
  }

  get solicitudesFiltradas(): any[] {
    const texto = this.filtroSolicitudes.toLowerCase().trim();
    return this.solicitudes.filter(s => {
      if (this.tabSolicitudes !== 'TODAS' && s.estado !== this.tabSolicitudes) return false;
      if (this.filtroOperacion && s.tipoOperacion !== this.filtroOperacion) return false;
      if (!texto) return true;
      return [
        '#' + s.idSolicitud, s.idSolicitud, this.nombreBienSolicitud(s), this.claveBienSolicitud(s),
        s.nombreSolicitante, s.correoSolicitante, s.nombreSede, this.etiquetaOperacion(s.tipoOperacion)
      ].join(' ').toLowerCase().includes(texto);
    });
  }

  get mensajeVacioSolicitudes(): string {
    if (this.filtroSolicitudes || this.filtroOperacion) return 'Ninguna solicitud coincide con los filtros';
    if (this.tabSolicitudes === 'PENDIENTE') {
      return this.vistaSolicitudes === 'autorizacion' ? 'No hay solicitudes pendientes por revisar' : 'No tienes solicitudes en revisión';
    }
    if (this.tabSolicitudes === 'APROBADA') return 'No hay solicitudes aprobadas';
    if (this.tabSolicitudes === 'RECHAZADA') return 'No hay solicitudes rechazadas';
    return 'Aún no hay solicitudes';
  }

  conteoSolicitudes(estado: string): number {
    return this.solicitudes.filter(s => s.estado === estado).length;
  }

  cambiarVistaSolicitudes(vista: 'autorizacion' | 'mias'): void {
    if (this.vistaSolicitudes === vista) return;
    this.vistaSolicitudes = vista;
    this.tabSolicitudes = vista === 'autorizacion' ? 'PENDIENTE' : 'TODAS';
    this.solicitudes = [];
    this.loadSolicitudes();
  }

  loadSolicitudes(): void {
    if (!this.logged) return;
    if (!this.esRevisor) this.vistaSolicitudes = 'mias';

    let endpoint = '/solicitudes-bienes/mis-solicitudes';
    if (this.vistaSolicitudes === 'autorizacion') {
      // El ADMIN consulta el historial completo; un revisor por sede solo sus pendientes
      endpoint = this.isAdmin ? '/solicitudes-bienes' : '/solicitudes-bienes/pendientes';
    }

    this.solicitudesLoading = true;
    this.solicitudesError = '';
    this.http.get<any[]>(this.api + endpoint).subscribe({
      next: rows => {
        this.solicitudes = (rows || []).sort((a, b) =>
          String(b.fechaSolicitud || '').localeCompare(String(a.fechaSolicitud || '')));
        this.solicitudesLoading = false;
      },
      error: e => {
        this.solicitudes = [];
        this.solicitudesLoading = false;
        this.solicitudesError = this.mensajeError(e, 'No fue posible consultar las solicitudes.');
      }
    });
  }

  puedeResolver(s: any): boolean {
    if (!s || s.estado !== 'PENDIENTE' || this.vistaSolicitudes !== 'autorizacion') return false;
    if (this.isAdmin) return true;
    return this.canAuthorize && s.correoSolicitante !== this.email;
  }

  abrirSolicitud(s: any): void {
    this.solicitudSeleccionada = s;
    this.motivoResolucion = '';
    this.motivoRequerido = false;
    this.errorResolucion = '';
    this.resolviendo = '';
    this.modoRevocar = false;
    this.motivoRevocacion = '';
    this.revocacionRequerida = false;
    this.errorRevocacion = '';
    this.construirCamposSolicitud(s);
  }

  cerrarSolicitud(): void {
    if (this.resolviendo || this.revocando) return;
    this.solicitudSeleccionada = null;
  }

  iniciarRevocacion(): void {
    this.modoRevocar = true;
    this.motivoRevocacion = '';
    this.revocacionRequerida = false;
    this.errorRevocacion = '';
  }

  // El admin se equivocó: deshace lo aplicado y la solicitud regresa a pendientes
  revocarSolicitud(): void {
    const s = this.solicitudSeleccionada;
    if (!s || this.revocando) return;
    const motivo = (this.motivoRevocacion || '').trim();
    if (!motivo) { this.revocacionRequerida = true; return; }

    this.revocando = true;
    this.errorRevocacion = '';
    this.http.post<any>(this.api + '/solicitudes-bienes/' + s.idSolicitud + '/revocar', { motivo }).subscribe({
      next: () => {
        this.revocando = false;
        this.modoRevocar = false;
        this.solicitudSeleccionada = null;
        this.tabSolicitudes = 'PENDIENTE';
        this.notificar('pendiente', 'Autorización revocada',
          'La solicitud #' + s.idSolicitud + ' regresó a la bandeja como pendiente.' +
          (s.estado === 'APROBADA' ? ' El cambio se deshizo en el inventario.' : ''), 4500);
        this.loadSolicitudes();
        if (s.estado === 'APROBADA') this.loadBienes();
      },
      error: e => {
        this.revocando = false;
        this.errorRevocacion = this.mensajeError(e, 'No fue posible revocar la autorización.');
      }
    });
  }

  resolverSolicitud(aprobar: boolean): void {
    const s = this.solicitudSeleccionada;
    if (!s || this.resolviendo) return;

    const motivo = (this.motivoResolucion || '').trim();
    if (!aprobar && !motivo) {
      this.motivoRequerido = true;
      return;
    }

    this.resolviendo = aprobar ? 'aprobar' : 'rechazar';
    this.errorResolucion = '';
    const bienPrevio = s.idBien ? this.bienes.find(b => b.id_bien === s.idBien) : null;
    const accion = aprobar ? 'aprobar' : 'rechazar';

    this.http.post<any>(
      this.api + '/solicitudes-bienes/' + s.idSolicitud + '/' + accion,
      { motivoResolucion: motivo || null }
    ).subscribe({
      next: r => {
        this.resolviendo = '';
        this.solicitudSeleccionada = null;
        if (aprobar) {
          const conFormato = s.tipoOperacion === 'ALTA' || s.tipoOperacion === 'BAJA';
          const resuelta = { ...s, estado: 'APROBADA', nombreRevisor: this.miNombre, motivoResolucion: motivo || null };
          this.notificar('exito', 'Solicitud aprobada',
            'La solicitud #' + s.idSolicitud + ' se aplicó al inventario.' + (conFormato ? ' Su formato quedó guardado en Documentos.' : ''), 3500,
            conFormato ? { texto: '📄 Ver formato de ' + (s.tipoOperacion === 'ALTA' ? 'alta' : 'baja'),
                           fn: () => this.verOFormato(r?.idDocumento, () => this.formatoDesdeSolicitud(resuelta, bienPrevio)) } : undefined);
        }
        else this.notificar('rechazo', 'Solicitud rechazada', 'Le avisamos al solicitante el motivo del rechazo.');
        this.loadSolicitudes();
        if (aprobar) this.loadBienes();
      },
      error: e => {
        this.resolviendo = '';
        this.errorResolucion = this.mensajeError(e, aprobar
          ? 'No fue posible aprobar la solicitud.'
          : 'No fue posible rechazar la solicitud.');
        if (e.status === 409) this.loadSolicitudes();
      }
    });
  }

  private construirCamposSolicitud(s: any): void {
    const datos = s.datosPropuestos || {};
    const actual = s.idBien ? this.bienes.find(b => b.id_bien === s.idBien) : null;
    this.hayValoresActuales = !!actual;

    this.camposSolicitud = this.camposBien
      .filter(c => s.tipoOperacion !== 'ALTA' || this.tieneValor(datos[c.clave]) || c.siempre)
      .map(c => {
        const propuesto = this.valorLegible(c, datos[c.clave]);
        const valorActual = actual ? this.valorLegible(c, actual[c.claveActual]) : '—';
        return {
          etiqueta: c.etiqueta,
          propuesto,
          actual: valorActual,
          cambio: !!actual && propuesto !== valorActual
        };
      });
  }

  private tieneValor(v: any): boolean {
    return v !== null && v !== undefined && String(v).trim() !== '';
  }

  private valorLegible(campo: any, valor: any): string {
    if (typeof valor === 'boolean') return valor ? 'Sí' : 'No';
    if (!this.tieneValor(valor)) return '—';
    if (campo.catalogo) {
      const fila = (this.catalogs[campo.catalogo] || []).find((x: any) => String(x[campo.idCatalogo]) === String(valor));
      return fila ? String(fila[campo.nombreCatalogo]) : '#' + valor;
    }
    return String(valor).replace(/_/g, ' ');
  }

  nombreBienSolicitud(s: any): string {
    return s.articuloBien || s.datosPropuestos?.articulo || (s.idBien ? 'Bien #' + s.idBien : 'Nuevo bien');
  }

  claveBienSolicitud(s: any): string {
    return s.claveBien || s.datosPropuestos?.clave || (s.tipoOperacion === 'ALTA' ? 'Alta nueva' : 'Sin clave');
  }

  etiquetaOperacion(tipo: string): string {
    return ({ ALTA: 'Alta', MODIFICACION: 'Modificación', BAJA: 'Baja' } as any)[tipo] || tipo || '—';
  }

  etiquetaEstado(estado: string): string {
    return ({ PENDIENTE: 'Pendiente', APROBADA: 'Aprobada', RECHAZADA: 'Rechazada' } as any)[estado] || estado;
  }

  claseEstadoSolicitud(estado: string): string {
    return ({ PENDIENTE: 'status-warning', APROBADA: 'status-good', RECHAZADA: 'status-bad' } as any)[estado] || 'status-neutral';
  }

  descripcionOperacion(s: any): string {
    const bien = this.nombreBienSolicitud(s);
    if (s.tipoOperacion === 'ALTA') return 'Solicitud para registrar el bien "' + bien + '" en el inventario.';
    if (s.tipoOperacion === 'MODIFICACION') return 'Solicitud para modificar los datos del bien "' + bien + '".';
    if (s.tipoOperacion === 'BAJA') return 'Solicitud para dar de baja el bien "' + bien + '".';
    return 'Solicitud de inventario.';
  }

  iniciales(texto: string): string {
    const base = String(texto || '?').split('@')[0];
    return base.split(/[\s._-]+/).filter(Boolean).map(x => x.charAt(0)).join('').substring(0, 2).toUpperCase() || '?';
  }

  private mensajeError(e: any, porDefecto: string): string {
    return e?.error?.message || e?.error?.error || porDefecto;
  }

  // =========================
  // INVENTARIO
  // =========================

  // Descarga el inventario una sola vez; búsqueda, filtros, orden y paginación se hacen en el navegador
  loadBienes(): void {
    if (!this.logged) return;
    this.bienesCargando = true;
    this.bienesError = '';

    this.http.get<any[]>(this.api + '/bienes').subscribe({
      next: v => {
        this.bienes = (v || []).map(b => ({
          ...b,
          _busqueda: [b.clave, b.etiqueta, b.articulo, b.modelo, b.marca, b.numero_serie, b.responsable, b.usuario_directo,
                      b.ubicacion, b.sede, b.clasificacion, b.subcategoria, b.color]
            .filter(x => x !== null && x !== undefined).join(' ').toLowerCase()
        }));
        this.calcularResumenBienes();
        this.aplicarFiltrosBienes(false);
        // Mantiene el panel lateral sincronizado si estaba abierto
        if (this.bienSeleccionado) {
          this.bienSeleccionado = this.bienes.find(b => b.id_bien === this.bienSeleccionado.id_bien) || null;
        }
        this.bienesCargando = false;
      },
      error: e => {
        this.bienesCargando = false;
        this.bienesError = this.mensajeError(e, 'Revisa tu conexión e inténtalo de nuevo.');
      }
    });
  }

  private calcularResumenBienes(): void {
    let piezas = 0, buenos = 0, disponibles = 0;
    const porClasificacion = new Map<any, { id: any; nombre: string; total: number }>();

    for (const b of this.bienes) {
      piezas += Number(b.cantidad_piezas || 1);
      if (String(b.estado_fisico || '').toUpperCase() === 'BUENO') buenos++;
      if (String(b.estatus_asignacion || '').toUpperCase() === 'SIN_ASIGNAR') disponibles++;
      const id = b.id_clasificacion ?? 'sin';
      const grupo = porClasificacion.get(id) || { id: b.id_clasificacion ?? null, nombre: b.clasificacion || 'Sin clasificación', total: 0 };
      grupo.total++;
      porClasificacion.set(id, grupo);
    }

    this.totalPiezas = piezas;
    this.bienesBuenEstado = buenos;
    this.bienesDisponibles = disponibles;
    const total = this.bienes.length || 1;
    this.resumenClasificaciones = Array.from(porClasificacion.values())
      .filter(g => g.id !== null)
      .sort((a, b) => b.total - a.total)
      .map(g => ({ ...g, pct: Math.round((g.total / total) * 1000) / 10 }));
    this.coloresClasificacion = {};
    this.resumenClasificaciones.forEach((c, i) => this.coloresClasificacion[String(c.id)] = this.colorClasificacion(i));
    this.calcularSedesYPersonas();
  }

  // Agrupa los bienes por sede y por persona para las vistas de tarjetas
  private calcularSedesYPersonas(): void {
    const sedes = new Map<any, any>();
    const personas = new Map<any, any>();

    for (const b of this.bienes) {
      const idSede = b.id_sede ?? 'sin';
      const s = sedes.get(idSede) || { id: b.id_sede ?? null, nombre: b.sede || 'Sin sede', total: 0, piezas: 0, buenos: 0,
                                         ubic: new Set<any>(), pers: new Set<any>(), clas: new Map<any, any>() };
      s.total++;
      s.piezas += Number(b.cantidad_piezas || 1);
      if (String(b.estado_fisico || '').toUpperCase() === 'BUENO') s.buenos++;
      if (b.id_ubicacion) s.ubic.add(b.id_ubicacion);
      if (b.id_responsable_area) s.pers.add(b.id_responsable_area);
      if (b.id_usuario_directo) s.pers.add(b.id_usuario_directo);
      if (b.id_clasificacion) {
        const c = s.clas.get(b.id_clasificacion) || { id: b.id_clasificacion, nombre: b.clasificacion || '', total: 0 };
        c.total++;
        s.clas.set(b.id_clasificacion, c);
      }
      sedes.set(idSede, s);

      const asignar = (id: any, nombre: string, rol: 'responsable' | 'usuario') => {
        if (!id) return;
        const p = personas.get(id) || { id, nombre: nombre || 'Sin nombre', total: 0, responsable: 0, usuario: 0, ids: new Set<any>(), sedesSet: new Set<string>() };
        p[rol]++;
        if (!p.ids.has(b.id_bien)) { p.ids.add(b.id_bien); p.total++; }
        if (b.sede) p.sedesSet.add(b.sede);
        personas.set(id, p);
      };
      asignar(b.id_responsable_area, b.responsable, 'responsable');
      asignar(b.id_usuario_directo, b.usuario_directo, 'usuario');
    }

    this.resumenSedes = Array.from(sedes.values())
      .map(s => ({
        id: s.id, nombre: s.nombre, total: s.total, piezas: s.piezas,
        pctBueno: s.total ? Math.round((s.buenos / s.total) * 100) : 0,
        ubicaciones: s.ubic.size, personas: s.pers.size,
        top: Array.from(s.clas.values()).sort((a: any, b: any) => b.total - a.total).slice(0, 3)
      }))
      .sort((a, b) => b.total - a.total);

    this.resumenPersonas = Array.from(personas.values())
      .map(p => ({ id: p.id, nombre: p.nombre, total: p.total, responsable: p.responsable, usuario: p.usuario,
                   sedes: Array.from(p.sedesSet).sort().join(', '), _busqueda: String(p.nombre || '').toLowerCase() }))
      .sort((a, b) => b.total - a.total || a.nombre.localeCompare(b.nombre, 'es'));
  }

  get personasFiltradas(): any[] {
    const palabras = this.busquedaPersonas.toLowerCase().trim().split(/\s+/).filter(Boolean);
    return palabras.length
      ? this.resumenPersonas.filter(p => palabras.every(w => p._busqueda.includes(w)))
      : this.resumenPersonas;
  }

  colorDeClasificacion(id: any): string {
    return this.coloresClasificacion[String(id)] || '#c3ccd8';
  }

  // Navegación entre las vistas del inventario
  abrirVistaInv(vista: 'sedes' | 'bienes' | 'personas' | 'completo'): void {
    this.origenVistaInv = 'menu';
    if (vista === 'bienes' || vista === 'completo') {
      this.filtroPersona = null;
      if (this.isAdmin) this.filtroSede = null;
      this.aplicarFiltrosBienes();
    }
    if (vista === 'personas') { this.busquedaPersonas = ''; this.personasVisibles = 48; }
    this.vistaInv = vista;
    this.subirAlInicio();
  }

  verBienesDeSede(s: any): void {
    this.filtroPersona = null;
    this.filtroSede = s.id;
    this.origenVistaInv = 'sedes';
    this.aplicarFiltrosBienes();
    this.vistaInv = 'bienes';
    this.subirAlInicio();
  }

  verBienesDePersona(p: any): void {
    if (this.isAdmin) this.filtroSede = null;
    this.filtroPersona = p.id;
    this.origenVistaInv = 'personas';
    this.aplicarFiltrosBienes();
    this.vistaInv = 'bienes';
    this.subirAlInicio();
  }

  volverInventario(): void {
    if (this.vistaInv === 'bienes' && this.origenVistaInv !== 'menu') {
      this.vistaInv = this.origenVistaInv;
      this.origenVistaInv = 'menu';
      this.filtroPersona = null;
      if (this.isAdmin) this.filtroSede = null;
      this.aplicarFiltrosBienes();
    } else {
      this.vistaInv = 'menu';
    }
    this.subirAlInicio();
  }

  get textoVolverInv(): string {
    if (this.vistaInv === 'bienes' && this.origenVistaInv === 'sedes') return 'Sedes';
    if (this.vistaInv === 'bienes' && this.origenVistaInv === 'personas') return 'Personas';
    return 'Bienes';
  }

  get tituloVistaInv(): string {
    switch (this.vistaInv) {
      case 'sedes': return 'Sedes';
      case 'personas': return 'Personas';
      case 'completo': return 'Inventario completo';
      case 'bienes': {
        if (this.filtroPersona) return this.resumenPersonas.find(p => p.id === this.filtroPersona)?.nombre || 'Bienes';
        if (this.filtroSede && this.origenVistaInv === 'sedes') return this.resumenSedes.find(s => s.id === this.filtroSede)?.nombre || 'Bienes';
        return 'Bienes';
      }
      default: return 'Bienes';
    }
  }

  get subtituloVistaInv(): string {
    switch (this.vistaInv) {
      case 'sedes': return 'Selecciona una sede para ver sus bienes.';
      case 'personas': return 'Selecciona una persona para ver los bienes que tiene a su cargo.';
      case 'completo': return 'Todos los bienes en tabla, con búsqueda, filtros y columnas a tu medida.';
      default: return this.filtroPersona ? 'Bienes de los que es responsable o usuario directo.' : 'Toca una tarjeta para ver el detalle del bien.';
    }
  }

  private subirAlInicio(): void {
    try { document.querySelector('.content-area')?.scrollTo({ top: 0 }); window.scrollTo({ top: 0 }); } catch { /* opcional */ }
  }

  // Búsqueda con pausa para no recalcular en cada tecla
  buscarBienes(): void {
    clearTimeout(this.busquedaTimer);
    this.busquedaTimer = setTimeout(() => this.aplicarFiltrosBienes(), 250);
  }

  aplicarFiltrosBienes(reiniciarPagina = true): void {
    const texto = this.filter.toLowerCase().trim();
    const palabras = texto ? texto.split(/\s+/) : [];

    let lista = this.bienes.filter(b =>
      (!this.statusFilter || b.estatus_inventario === this.statusFilter) &&
      (!this.filtroSede || b.id_sede === this.filtroSede) &&
      (!this.filtroClasificacion || b.id_clasificacion === this.filtroClasificacion) &&
      (!this.filtroEstadoFisico || String(b.estado_fisico || '').toUpperCase() === this.filtroEstadoFisico) &&
      (!this.filtroAsignacion || String(b.estatus_asignacion || '').toUpperCase() === this.filtroAsignacion) &&
      (!this.filtroCM || this.coincideCM(b, this.filtroCM)) &&
      (!this.filtroPersona || b.id_responsable_area === this.filtroPersona || b.id_usuario_directo === this.filtroPersona) &&
      palabras.every(p => b._busqueda.includes(p))
    );

    const { key, dir } = this.ordenBienes;
    if (key) {
      const factor = dir === 'asc' ? 1 : -1;
      const campo = key === 'articulo' ? 'articulo' : key;
      lista = [...lista].sort((a, b) => {
        const va = a[campo], vb = b[campo];
        if (va === vb) return 0;
        if (va === null || va === undefined || va === '') return 1;
        if (vb === null || vb === undefined || vb === '') return -1;
        return (typeof va === 'number' && typeof vb === 'number'
          ? va - vb
          : String(va).localeCompare(String(vb), 'es', { numeric: true })) * factor;
      });
    }

    this.bienesVista = lista;
    if (reiniciarPagina) { this.paginaActualBienes = 1; this.tarjetasVisibles = 36; }
    else this.paginaActualBienes = Math.min(this.paginaActualBienes, this.totalPaginasBienes);
  }

  ordenarBienes(key: string): void {
    this.ordenBienes = this.ordenBienes.key === key
      ? { key, dir: this.ordenBienes.dir === 'asc' ? 'desc' : 'asc' }
      : { key, dir: 'asc' };
    this.aplicarFiltrosBienes(false);
  }

  setFiltroClasificacion(id: any): void {
    this.filtroClasificacion = id;
    this.aplicarFiltrosBienes();
  }

  // ----- Etiquetas QR en blanco -----

  etq: any = null;
  etqAsignar: any = null;
  mathMax = Math.max;
  mathMin = Math.min;

  codigoEtiqueta(n: number): string {
    return 'DZ-' + String(Math.max(0, n || 0)).padStart(6, '0');
  }

  // El QR lleva solo la dirección de DOZZER con el código (sin datos del bien)
  urlEtiqueta(codigo: string): string {
    return window.location.origin + '/?e=' + codigo;
  }

  // "…/?e=DZ-000123", "DZ-000123" o "dz123" → "DZ-000123"
  codigoDeTexto(t: string): string {
    let v = String(t || '').trim();
    try { if (/^https?:\/\//i.test(v)) v = new URL(v).searchParams.get('e') || ''; } catch { v = ''; }
    const m = v.toUpperCase().replace(/[^A-Z0-9]/g, '').match(/^DZ(\d{1,9})$/);
    return m ? this.codigoEtiqueta(parseInt(m[1], 10)) : '';
  }

  hojasPara(n: number): number {
    return Math.ceil(Math.max(0, n || 0) / 30);
  }

  get etqFaltan(): number {
    return (this.etq?.resumen?.porSede || []).reduce((t: number, x: any) => t + (x.sin_etiqueta || 0), 0);
  }

  // Lo que falta por pegar menos lo ya impreso sin pegar, con 5% extra por etiquetas dañadas
  get hojasSugeridas(): number {
    const r = this.etq?.resumen;
    if (!r) return 0;
    return this.hojasPara(Math.max(0, this.etqFaltan - (r.libres || 0)) * 1.05);
  }

  abrirEtiquetas(): void {
    this.etq = { cargando: true, generando: false, error: '', resumen: null, hojas: 1, verSedes: false };
    this.http.get<any>(this.api + '/etiquetas/resumen').subscribe({
      next: r => {
        this.etq.resumen = r;
        this.etq.cargando = false;
        this.etq.hojas = Math.max(1, Math.min(30, this.hojasSugeridas));
      },
      error: e => { this.etq.cargando = false; this.etq.error = this.mensajeError(e, 'No se pudo calcular.'); }
    });
  }

  generarEtiquetas(): void {
    const e = this.etq;
    const hojas = Math.floor(Number(e.hojas));
    if (!(hojas >= 1 && hojas <= 30)) { e.error = 'Imprime entre 1 y 30 hojas por vez.'; return; }
    const ventana = window.open('', '_blank');
    if (!ventana) { e.error = 'El navegador bloqueó la ventana de impresión. Permite ventanas emergentes para DOZZER.'; return; }
    ventana.document.write('<p style="font-family:Arial;padding:30px;color:#41546b">Generando etiquetas...</p>');
    e.generando = true; e.error = '';
    this.http.post<any>(this.api + '/etiquetas/lote', { cantidad: hojas * 30 }).subscribe({
      next: async r => {
        await this.imprimirCodigos(r.codigos || [], ventana);
        e.generando = false;
        this.abrirEtiquetas();
      },
      error: err => { ventana.close(); e.generando = false; e.error = this.mensajeError(err, 'No se generaron las etiquetas.'); }
    });
  }

  reimprimirLibres(): void {
    const ventana = window.open('', '_blank');
    if (!ventana) { this.etq.error = 'El navegador bloqueó la ventana de impresión.'; return; }
    this.etq.generando = true;
    this.http.get<string[]>(this.api + '/etiquetas/libres', { params: { limite: 900 } }).subscribe({
      next: async cods => { await this.imprimirCodigos(cods || [], ventana); this.etq.generando = false; },
      error: err => { ventana.close(); this.etq.generando = false; this.etq.error = this.mensajeError(err, 'No se pudo.'); }
    });
  }

  // Hoja carta de 30 etiquetas (3 × 10, 2.625" × 1", tipo Avery 5160)
  async imprimirCodigos(codigos: string[], ventanaAbierta?: Window | null): Promise<void> {
    if (!codigos.length) return;
    const ventana = ventanaAbierta || window.open('', '_blank');
    if (!ventana) { this.notificar('error', 'Ventana bloqueada', 'Permite ventanas emergentes para DOZZER e inténtalo de nuevo.', 0); return; }
    const e = (v: any) => String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' } as any)[c]);
    const etiquetas: string[] = [];
    for (let i = 0; i < codigos.length; i++) {
      const svg = await QRCode.toString(this.urlEtiqueta(codigos[i]), { type: 'svg', errorCorrectionLevel: 'M', margin: 0, color: { dark: '#000000' } });
      const pos = i % 30, col = pos % 3, fila = Math.floor(pos / 3);
      etiquetas.push((pos === 0 && i > 0 ? '</div><div class="hoja">' : '') +
        '<div class="et" style="left:' + (0.1875 + col * 2.75) + 'in;top:' + (0.5 + fila) + 'in">' +
        '<div class="qr">' + svg + '</div><div class="tx"><div class="mk">DOZZER</div><div class="cod">' + e(codigos[i]) + '</div>' +
        '<div class="pie">Propiedad de DOZZER Construcción y Asesoría</div></div></div>');
    }
    ventana.document.open();
    ventana.document.write('<!doctype html><html lang="es"><head><meta charset="utf-8"><title>Etiquetas DOZZER ' + e(codigos[0]) + ' a ' + e(codigos[codigos.length - 1]) + '</title><style>' +
      '@page{size:letter;margin:0}*{box-sizing:border-box}body{margin:0;font-family:Arial,Helvetica,sans-serif;color:#111;-webkit-print-color-adjust:exact;print-color-adjust:exact}' +
      '.hoja{position:relative;width:8.5in;height:11in;page-break-after:always;overflow:hidden}.hoja:last-child{page-break-after:auto}' +
      '.et{position:absolute;width:2.625in;height:1in;padding:0.08in 0.1in;display:flex;align-items:center;gap:0.09in}' +
      '.qr{width:0.84in;height:0.84in;flex-shrink:0}.qr svg{width:100%;height:100%;display:block}' +
      '.tx{min-width:0;display:flex;flex-direction:column;gap:0.02in}.mk{font-size:7.5pt;font-weight:800;letter-spacing:1.4pt;color:#17263d}' +
      '.cod{font-family:"Courier New",monospace;font-size:13.5pt;font-weight:800;letter-spacing:0.3pt}.pie{font-size:5pt;color:#555;line-height:1.2}' +
      '.barra{position:sticky;top:0;z-index:2;background:#17263d;color:#fff;padding:10px 16px;font-size:13px;display:flex;justify-content:space-between;align-items:center}' +
      '.barra button{background:#fff;color:#17263d;border:0;border-radius:6px;padding:7px 14px;font-weight:700;cursor:pointer}@media print{.barra{display:none}}' +
      '</style></head><body><div class="barra"><span>' + codigos.length + ' etiquetas · ' + Math.ceil(codigos.length / 30) + ' hoja(s) · ' + e(codigos[0]) + ' a ' + e(codigos[codigos.length - 1]) +
      '</span><button onclick="window.print()">Imprimir</button></div><div class="hoja">' + etiquetas.join('') + '</div></body></html>');
    ventana.document.close();
    ventana.focus();
    setTimeout(() => ventana.print(), 400);
  }

  abrirAsignarEtiqueta(b: any): void {
    this.etqAsignar = { bien: b, codigo: '', reemplazar: false, guardando: false, error: '' };
  }

  async leerEtiquetaDeFoto(ev: Event): Promise<void> {
    const input = ev.target as HTMLInputElement;
    const archivo = input.files?.[0];
    input.value = '';
    if (!archivo || !this.etqAsignar) return;
    try {
      const img = await this.cargarImagen(archivo);
      const escala = Math.min(1, 1400 / Math.max(img.width, img.height));
      const canvas = document.createElement('canvas');
      canvas.width = Math.round(img.width * escala); canvas.height = Math.round(img.height * escala);
      const ctx = canvas.getContext('2d') as CanvasRenderingContext2D;
      ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
      const qr = jsQR(ctx.getImageData(0, 0, canvas.width, canvas.height).data, canvas.width, canvas.height);
      const cod = qr?.data ? this.codigoDeTexto(qr.data) : '';
      if (cod) { this.etqAsignar.codigo = cod; this.etqAsignar.error = ''; }
      else this.etqAsignar.error = 'No se encontró un QR de etiqueta DOZZER en la foto.';
    } catch { this.etqAsignar.error = 'No se pudo leer la foto.'; }
  }

  confirmarAsignarEtiqueta(): void {
    const a = this.etqAsignar;
    const cod = this.codigoDeTexto(a?.codigo);
    if (!cod) { a.error = 'Escribe un código válido, por ejemplo DZ-000123.'; return; }
    a.guardando = true; a.error = '';
    this.http.post<any>(this.api + '/etiquetas/' + encodeURIComponent(cod) + '/asignar', { idBien: a.bien.id_bien, reemplazar: !!a.reemplazar }).subscribe({
      next: () => {
        a.bien.etiqueta = cod;
        const enLista = this.bienes.find(x => x.id_bien === a.bien.id_bien);
        if (enLista) enLista.etiqueta = cod;
        this.etqAsignar = null;
        if (this.bienSeleccionado?.id_bien === a.bien.id_bien) this.abrirBien(this.bienSeleccionado);
        this.notificar('exito', 'Etiqueta asignada', cod + ' → ' + a.bien.clave, 2500);
      },
      error: e => { a.guardando = false; a.error = this.mensajeError(e, 'No se pudo asignar.'); }
    });
  }

  // Se abrió DOZZER escaneando una etiqueta con la cámara del celular (…/?e=DZ-000123)
  private abrirEtiquetaDesdeUrl(codigo: string): void {
    this.http.get<any>(this.api + '/etiquetas/' + encodeURIComponent(codigo)).subscribe({
      next: r => {
        if (r.estado === 'ASIGNADA' && r.idBien) {
          this.section = 'bienes';
          const b = this.bienes.find(x => x.id_bien === r.idBien);
          if (b) this.abrirBien(b);
          else this.http.get<any>(this.api + '/bienes/' + r.idBien).subscribe({ next: x => this.abrirBien(x), error: () => {} });
          return;
        }
        const msg: any = { LIBRE: 'Esta etiqueta todavía no está pegada en ningún bien. Asígnala desde la ficha del bien o en el levantamiento.',
          ANULADA: 'Esta etiqueta fue anulada.', NO_EXISTE: 'Este código no es una etiqueta de DOZZER.' };
        this.notificar('info', 'Etiqueta ' + codigo, r.estado === 'ASIGNADA' ? 'Es de un bien de otra sede.' : (msg[r.estado] || ''), 0);
      },
      error: () => {}
    });
  }

  // ----- Calibración y mantenimiento -----

  // vencida | proxima (30 días) | ok | sin (sin fecha)
  estadoFechaCM(f: any): 'vencida' | 'proxima' | 'ok' | 'sin' {
    if (!f) return 'sin';
    const hoy = new Date(); hoy.setHours(0, 0, 0, 0);
    const d = new Date(String(f).substring(0, 10) + 'T00:00:00');
    const dias = Math.round((d.getTime() - hoy.getTime()) / 86400000);
    return dias < 0 ? 'vencida' : dias <= 30 ? 'proxima' : 'ok';
  }

  textoFechaCM(f: any): string {
    const e = this.estadoFechaCM(f);
    if (e === 'sin') return 'sin fecha';
    const d = new Date(String(f).substring(0, 10) + 'T00:00:00');
    const txt = d.toLocaleDateString('es-MX', { day: '2-digit', month: '2-digit', year: 'numeric' });
    return e === 'vencida' ? 'vencida el ' + txt : 'próxima: ' + txt;
  }

  private coincideCM(b: any, filtro: string): boolean {
    const cal = !!b.requiere_calibracion, man = !!b.requiere_mantenimiento;
    switch (filtro) {
      case 'calibracion': return cal;
      case 'mantenimiento': return man;
      case 'sinfecha': return (cal && !b.fecha_calibracion) || (man && !b.fecha_mantenimiento);
      case 'atencion': return (cal && ['vencida', 'proxima'].includes(this.estadoFechaCM(b.fecha_calibracion)))
        || (man && ['vencida', 'proxima'].includes(this.estadoFechaCM(b.fecha_mantenimiento)));
      default: return true;
    }
  }

  // Resumen para el aviso de Inicio (solo bienes activos)
  get alertasCM(): { vencidas: number; proximas: number; sinFecha: number; total: number } {
    let vencidas = 0, proximas = 0, sinFecha = 0, total = 0;
    for (const b of this.bienes) {
      for (const [req, f] of [[b.requiere_calibracion, b.fecha_calibracion], [b.requiere_mantenimiento, b.fecha_mantenimiento]]) {
        if (!req) continue;
        total++;
        const e = this.estadoFechaCM(f);
        if (e === 'vencida') vencidas++; else if (e === 'proxima') proximas++; else if (e === 'sin') sinFecha++;
      }
    }
    return { vencidas, proximas, sinFecha, total };
  }

  verAlertasCM(filtro: 'atencion' | 'sinfecha'): void {
    this.section = 'bienes';
    this.abrirVistaInv('completo');
    this.limpiarFiltrosBienes();
    this.filtroCM = filtro;
    this.aplicarFiltrosBienes();
  }

  limpiarFiltrosBienes(): void {
    this.filter = '';
    this.statusFilter = '';
    this.filtroSede = null;
    this.filtroClasificacion = null;
    this.filtroEstadoFisico = '';
    this.filtroAsignacion = '';
    this.filtroPersona = null;
    this.filtroCM = '';
    this.aplicarFiltrosBienes();
  }

  // Columnas visibles (se recuerdan en este navegador)
  columnaVisible(key: string): boolean {
    return this.columnasElegidas.includes(key);
  }

  toggleColumna(key: string): void {
    const col = this.columnasDisponibles.find(c => c.key === key);
    if (col?.fija) return;
    this.columnasElegidas = this.columnaVisible(key)
      ? this.columnasElegidas.filter(k => k !== key)
      : [...this.columnasElegidas, key];
    this.guardarColumnas();
  }

  aplicarVistaColumnas(vista: 'basica' | 'ubicacion' | 'completa'): void {
    const vistas: any = {
      basica: this.columnasPorDefecto(),
      ubicacion: ['clave', 'articulo', 'sede', 'ubicacion', 'responsable', 'usuario_directo', 'estatus_asignacion', 'estatus_inventario'],
      completa: this.columnasDisponibles.map(c => c.key)
    };
    this.columnasElegidas = vistas[vista];
    this.guardarColumnas();
  }

  private columnasPorDefecto(): string[] {
    return ['clave', 'articulo', ...(this.isAdmin ? ['sede'] : []), 'ubicacion', 'responsable', 'estatus_inventario'];
  }

  private cargarColumnas(): string[] {
    try {
      const guardadas = JSON.parse(localStorage.getItem('dozzer_columnas_inventario') || 'null');
      if (Array.isArray(guardadas) && guardadas.length) return Array.from(new Set(['clave', 'articulo', ...guardadas]));
    } catch { /* sin preferencias guardadas */ }
    return this.columnasPorDefecto();
  }

  private guardarColumnas(): void {
    try { localStorage.setItem('dozzer_columnas_inventario', JSON.stringify(this.columnasElegidas)); } catch { /* opcional */ }
  }

  abrirBien(b: any): void {
    this.bienSeleccionado = b;
    this.qrPreview = '';
    this.refrescarFotoBien(b);
    if (!b?.etiqueta) return;
    QRCode.toDataURL(this.urlEtiqueta(b.etiqueta), { errorCorrectionLevel: 'M', margin: 1, width: 220, color: { dark: '#17263d' } })
      .then(url => { if (this.bienSeleccionado?.id_bien === b.id_bien) this.qrPreview = url; })
      .catch(() => this.qrPreview = '');
  }

  // =========================
  // ETIQUETAS QR
  // =========================

  // Texto que va dentro del QR: los datos del bien al momento de imprimir, legibles por cualquier
  // lector de QR sin internet ni sesión. La versión compacta lleva menos datos para que el QR
  // siga siendo fácil de leer en etiquetas pequeñas.
  contenidoQr(b: any, compacto = false): string {
    const v = (x: any) => {
      const s = String(x ?? '').replace(/\s+/g, ' ').trim();
      return s.length > 60 ? s.substring(0, 57) + '...' : s;
    };

    // Con ficha publicada: enlace a la página, con los datos después del "#" (no se envían al servidor)
    if (QR_FICHA_URL) {
      const campos: [string, any][] = compacto
        ? [['c', b.clave], ['b', b.articulo], ['s', b.numero_serie], ['d', b.sede]]
        : [
            ['c', b.clave], ['b', b.articulo],
            ['m', [b.marca, b.modelo].filter(x => x && String(x).trim()).join(' ')],
            ['s', b.numero_serie], ['k', b.clasificacion], ['d', b.sede],
            ['u', b.ubicacion], ['r', b.responsable]
          ];
      const params = campos.filter(([, x]) => v(x))
        .map(([k, x]) => k + '=' + encodeURIComponent(v(x)))
        .concat('i=' + this.idCortoQr(b.codigo_qr));
      return QR_FICHA_URL + '#' + params.join('&');
    }
    const lineas: [string, any][] = compacto
      ? [['Clave', b.clave], ['Bien', b.articulo], ['Serie', b.numero_serie], ['Sede', b.sede]]
      : [
          ['Clave', b.clave], ['Bien', b.articulo],
          ['Marca/Modelo', [b.marca, b.modelo].filter(x => x && String(x).trim()).join(' ')],
          ['Serie', b.numero_serie], ['Clasificación', b.clasificacion],
          ['Sede', b.sede], ['Ubicación', b.ubicacion], ['Responsable', b.responsable]
        ];
    return ['DOZZER · INVENTARIO',
      ...lineas.filter(([, x]) => v(x)).map(([k, x]) => k + ': ' + v(x)),
      'ID: ' + this.idCortoQr(b.codigo_qr)
    ].join('\n');
  }

  idCortoQr(codigo: string): string {
    return codigo ? codigo.replace(/-/g, '').substring(0, 8).toUpperCase() : '—';
  }

  ubicacionEtiqueta(b: any): string {
    return [b.sede, b.ubicacion].filter(x => x && String(x).trim()).join(' · ');
  }

  // La ficha apunta a una dirección local (localhost o red privada): sirve para probar, no para imprimir en serio
  get fichaQrDePrueba(): boolean {
    return /^https?:\/\/(localhost|127\.|10\.|192\.168\.|172\.(1[6-9]|2\d|3[01])\.)/i.test(QR_FICHA_URL);
  }

  get fichaQrHost(): string {
    try { return new URL(QR_FICHA_URL).host; } catch { return QR_FICHA_URL; }
  }

  get etiquetasSeleccion(): any[] {
    const lista = this.etiquetasAlcance === 'pagina' ? this.paginaBienes : this.bienesVista;
    return lista.filter(b => b.codigo_qr);
  }

  get hojasEtiquetas(): number {
    const porHoja = this.etiquetasFormato === 'compacta' ? 24 : 10;
    return Math.ceil(this.etiquetasSeleccion.length / porHoja);
  }

  abrirModalEtiquetas(): void {
    this.modalEtiquetas = true;
    this.generarQrMuestra();
  }

  generarQrMuestra(): void {
    this.qrMuestra = '';
    setTimeout(() => {
      const primero = this.etiquetasSeleccion[0];
      if (!primero) return;
      QRCode.toDataURL(this.contenidoQr(primero, this.etiquetasFormato === 'compacta'), { errorCorrectionLevel: 'M', margin: 0, width: 200, color: { dark: '#17263d' } })
        .then(url => this.qrMuestra = url)
        .catch(() => this.qrMuestra = '');
    });
  }

  // Genera una hoja de impresión en una ventana nueva (se abre antes para que el navegador no la bloquee)
  async imprimirEtiquetas(lista: any[]): Promise<void> {
    const bienes = (lista || []).filter(b => b.codigo_qr);
    if (!bienes.length || this.generandoEtiquetas) return;

    const ventana = window.open('', '_blank');
    if (!ventana) {
      this.notificar('error', 'Ventana bloqueada', 'Tu navegador bloqueó la impresión. Permite las ventanas emergentes para DOZZER e inténtalo de nuevo.', 0);
      return;
    }
    ventana.document.write('<p style="font-family:Arial;padding:30px;color:#41546b">Generando ' + bienes.length + ' etiqueta(s)...</p>');

    this.generandoEtiquetas = true;
    try {
      const compacta = bienes.length > 1 && this.etiquetasFormato === 'compacta';
      const etiquetas: string[] = [];
      for (const b of bienes) {
        const svg = await QRCode.toString(this.contenidoQr(b, compacta), { type: 'svg', errorCorrectionLevel: 'M', margin: 0, color: { dark: '#17263d' } });
        etiquetas.push(this.htmlEtiqueta(b, svg, compacta));
      }
      ventana.document.open();
      ventana.document.write(this.htmlHojaEtiquetas(etiquetas, compacta, bienes.length === 1));
      ventana.document.close();
      ventana.focus();
      setTimeout(() => ventana.print(), 350);
      this.modalEtiquetas = false;
    } catch {
      ventana.close();
      this.notificar('error', 'No se generaron las etiquetas', 'Ocurrió un problema al preparar las etiquetas QR. Inténtalo de nuevo.', 0);
    } finally {
      this.generandoEtiquetas = false;
    }
  }

  private htmlEtiqueta(b: any, svg: string, compacta: boolean): string {
    const e = (v: any) => String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' } as any)[c]);
    const ubicacion = this.ubicacionEtiqueta(b);
    return '<div class="label">' +
      '<div class="qr">' + svg + '</div>' +
      '<div class="info">' +
        '<div class="brand"><span class="logo">D</span>DOZZER</div>' +
        '<div class="clave">' + e(b.clave) + '</div>' +
        '<div class="articulo">' + e(b.articulo) + '</div>' +
        (!compacta && ubicacion ? '<div class="meta">' + e(ubicacion) + '</div>' : '') +
        '<div class="id">ID ' + e(this.idCortoQr(b.codigo_qr)) + '</div>' +
        (!compacta ? '<div class="foot">Propiedad de DOZZER Construcción y Asesoría S.A. de C.V.</div>' : '') +
      '</div></div>';
  }

  private htmlHojaEtiquetas(etiquetas: string[], compacta: boolean, individual: boolean): string {
    const columnas = compacta ? 3 : 2;
    const ancho = compacta ? '64mm' : '95mm';
    const alto = compacta ? '33mm' : '50mm';
    const qr = compacta ? '25mm' : '38mm';
    return '<!doctype html><html lang="es"><head><meta charset="utf-8"><title>Etiquetas QR · DOZZER</title><style>' +
      '@page{size:letter;margin:' + (compacta ? '10mm 6mm' : '12mm 10mm') + '}' +
      '*{box-sizing:border-box}body{margin:0;font-family:Arial,Helvetica,sans-serif;color:#17263d;-webkit-print-color-adjust:exact;print-color-adjust:exact}' +
      '.sheet{display:grid;grid-template-columns:repeat(' + (individual ? 1 : columnas) + ',' + ancho + ');gap:' + (compacta ? '2mm 3mm' : '3mm 5mm') + ';justify-content:center}' +
      '.label{width:' + ancho + ';height:' + alto + ';border:0.3mm solid #c9d3df;border-radius:2.5mm;padding:' + (compacta ? '2.5mm' : '4mm') + ';display:flex;gap:' + (compacta ? '2.5mm' : '4mm') + ';align-items:center;overflow:hidden;break-inside:avoid;page-break-inside:avoid}' +
      '.qr{width:' + qr + ';height:' + qr + ';flex-shrink:0}.qr svg{width:100%;height:100%;display:block}' +
      '.info{min-width:0;flex:1;display:flex;flex-direction:column;gap:' + (compacta ? '0.8mm' : '1.4mm') + '}' +
      '.brand{display:flex;align-items:center;gap:1.5mm;font-weight:800;font-size:' + (compacta ? '6.5pt' : '8pt') + ';letter-spacing:.8pt;color:#244c78}' +
      '.logo{width:' + (compacta ? '3.6mm' : '4.6mm') + ';height:' + (compacta ? '3.6mm' : '4.6mm') + ';border-radius:1mm;background:#17263d;color:#fff;display:inline-grid;place-items:center;font-size:' + (compacta ? '6pt' : '7.5pt') + '}' +
      '.clave{font-size:' + (compacta ? '10pt' : '14pt') + ';font-weight:800;letter-spacing:.3pt;margin-top:' + (compacta ? '0.5mm' : '1mm') + '}' +
      '.articulo{font-size:' + (compacta ? '6.5pt' : '8pt') + ';font-weight:700;color:#33465e;line-height:1.2;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;text-transform:uppercase}' +
      '.meta{font-size:6.5pt;color:#5d6f86;line-height:1.25;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}' +
      '.id{font-family:"Courier New",monospace;font-size:' + (compacta ? '6pt' : '7pt') + ';color:#5d6f86;letter-spacing:.6pt}' +
      '.foot{margin-top:auto;font-size:5pt;color:#8a97a8;border-top:0.2mm solid #e1e7ef;padding-top:1mm}' +
      '.toolbar{position:sticky;top:0;background:#17263d;color:#fff;padding:10px 16px;font-size:13px;display:flex;justify-content:space-between;align-items:center;margin-bottom:14px}' +
      '.toolbar button{background:#fff;color:#17263d;border:0;border-radius:6px;padding:7px 14px;font-weight:700;cursor:pointer}' +
      '@media print{.toolbar{display:none}}' +
      '</style></head><body>' +
      '<div class="toolbar"><span>' + etiquetas.length + ' etiqueta(s) lista(s) para imprimir</span><button onclick="window.print()">Imprimir</button></div>' +
      '<div class="sheet">' + etiquetas.join('') + '</div></body></html>';
  }


  editarDesdePanel(b: any): void {
    this.bienSeleccionado = null;
    this.editBien(b);
  }

  trackBien(_: number, b: any): any {
    return b.id_bien;
  }

  marcaModelo(b: any): string {
    return [b.marca, b.modelo].filter(x => x && String(x).trim()).join(' · ') || 'Sin marca ni modelo';
  }

  textoCampo(v: any): string {
    return v === null || v === undefined || String(v).trim() === '' ? '—' : String(v).replace(/_/g, ' ');
  }

  porcentaje(valor: number): number {
    return this.bienes.length ? Math.round((valor / this.bienes.length) * 100) : 0;
  }

  claseEstadoFisico(estado: string): string {
    return ({ BUENO: 'status-good', REGULAR: 'status-warning', MALO: 'status-bad' } as any)[String(estado || '').toUpperCase()] || 'status-neutral';
  }

  colorClasificacion(i: number): string {
    const paleta = ['#3d6fb0', '#2f9b78', '#c4862f', '#7a5fc0', '#c05a62', '#3b8ea5', '#7d8a3a', '#a0617f'];
    return paleta[i % paleta.length];
  }

  newBien(): void {
    this.bienOriginal = null;
    this.bien = {
      clave: '',
      articulo: '',
      cantidadPiezas: 1,
      estatusInventario: 'ACTIVO',
      estatusAsignacion: 'SIN_ASIGNAR',
      requiereCalibracion: false, fechaCalibracion: null,
      requiereMantenimiento: false, fechaMantenimiento: null
    };
    this.editing = true;
  }

  editBien(b: any): void {
    this.bienOriginal = b;
    this.bien = {
      idBien: b.id_bien,
      clave: b.clave,
      articulo: b.articulo,
      idClasificacion: b.id_clasificacion,
      idSubcategoria: b.id_subcategoria ?? null,
      idMarca: b.id_marca,
      modelo: b.modelo,
      numeroSerie: b.numero_serie,
      idColor: b.id_color,
      unidadMedida: b.unidad_medida,
      cantidadPiezas: b.cantidad_piezas,
      idResponsableArea: b.id_responsable_area,
      idUsuarioDirecto: b.id_usuario_directo,
      idUbicacion: b.id_ubicacion,
      estatusInventario: b.estatus_inventario,
      estatusAsignacion: b.estatus_asignacion,
      estadoFisico: b.estado_fisico,
      caracteristicas: b.caracteristicas,
      requiereCalibracion: !!b.requiere_calibracion,
      fechaCalibracion: b.fecha_calibracion || null,
      requiereMantenimiento: !!b.requiere_mantenimiento,
      fechaMantenimiento: b.fecha_mantenimiento || null
    };
    this.editing = true;
  }

  // Si cambia la clasificación, se limpia la subcategoría que ya no le corresponde
  onClasificacionBienChange(): void {
    const sub = (this.catalogs.subcategorias || []).find((s: any) => s.id_subcategoria === this.bien.idSubcategoria);
    if (sub && sub.id_clasificacion !== this.bien.idClasificacion) this.bien.idSubcategoria = null;
  }

  // Al elegir una subcategoría, se completa su clasificación
  onSubcategoriaBienChange(): void {
    const sub = (this.catalogs.subcategorias || []).find((s: any) => s.id_subcategoria === this.bien.idSubcategoria);
    if (sub && !this.bien.idClasificacion) this.bien.idClasificacion = sub.id_clasificacion;
  }

  // Registros activos de un catálogo, conservando el valor ya asignado aunque esté inactivo
  activos(lista: any[], actual: any, campoId: string): any[] {
    return (lista || []).filter(x => x.estatus !== false || x[campoId] === actual);
  }

  // =========================
  // DIÁLOGO DE CONFIRMACIÓN
  // =========================

  ejecutarAccionAviso(): void {
    const accion = this.aviso?.accion;
    this.cerrarAviso();
    accion?.fn();
  }

  // =========================
  // FORMATOS PDF DE ALTA / BAJA
  // =========================

  get miNombre(): string {
    return this.nombreUsuario || this.email || 'Administrador';
  }

  get puedeCompartir(): boolean {
    try {
      const nav: any = navigator;
      return !!(this.documento && nav.canShare && nav.canShare({ files: [new File([this.documento.blob], this.documento.nombre, { type: 'application/pdf' })] }));
    } catch { return false; }
  }

  bienPorId(id: any): any {
    return this.bienes.find(b => b.id_bien === id) || null;
  }

  folio(tipo: 'ALTA' | 'BAJA', idSolicitud: any, referencia?: any): string {
    const anio = String(new Date().getFullYear()).slice(-2);
    const pref = tipo === 'ALTA' ? 'A' : 'B';
    if (idSolicitud) return pref + '-' + anio + '-' + String(idSolicitud).padStart(4, '0');
    const ref = referencia === null || referencia === undefined || referencia === '' ? '0' : String(referencia);
    return pref + '-' + anio + '-D' + (/^\d+$/.test(ref) ? ref.padStart(4, '0') : ref);
  }

  private fechaTexto(valor?: any): string {
    const f = valor ? new Date(valor) : new Date();
    const fecha = isNaN(f.getTime()) ? new Date() : f;
    return fecha.toLocaleDateString('es-MX', { day: '2-digit', month: 'long', year: 'numeric' }).toUpperCase();
  }

  private nombreDe(lista: any[], idKey: string, nombreKey: string, id: any): string {
    if (id === null || id === undefined || id === '') return '';
    const x = (lista || []).find((i: any) => String(i[idKey]) === String(id));
    return x ? String(x[nombreKey] || '') : '';
  }

  private filaDe(clave: any, articulo: any, marca: any, modelo: any, serie: any) {
    return {
      inventario: String(clave || ''),
      descripcion: [articulo, marca].filter(x => x).join(' · '),
      modeloSerie: [modelo, serie].filter(x => x).join(' / ')
    };
  }

  // Desde un bien del inventario (llaves en minúsculas con guion bajo)
  formatoDesdeBien(tipo: 'ALTA' | 'BAJA', b: any, extra: Partial<FormatoDoc> = {}): FormatoDoc {
    return {
      tipo,
      folio: extra.folio || this.folio(tipo, null, b?.id_bien),
      sede: b?.sede || '',
      fecha: extra.fecha || this.fechaTexto(tipo === 'ALTA' ? b?.fecha_alta : null),
      filas: [this.filaDe(b?.clave, b?.articulo, b?.marca, b?.modelo, b?.numero_serie)],
      detalle: [
        { etiqueta: 'Clasificación', valor: b?.clasificacion || '' },
        { etiqueta: 'Ubicación', valor: [b?.ubicacion, b?.sede].filter((x: any) => x).join(' · ') },
        { etiqueta: 'Responsable', valor: b?.responsable || '' },
        { etiqueta: 'Estado físico', valor: b?.estado_fisico || '' },
        { etiqueta: 'Piezas', valor: String(b?.cantidad_piezas || 1) },
        { etiqueta: 'Usuario directo', valor: b?.usuario_directo || '' }
      ],
      observaciones: extra.observaciones ?? (b?.caracteristicas || ''),
      solicito: extra.solicito || '',
      autorizo: extra.autorizo || '',
      correoSugerido: extra.correoSugerido,
      contactos: [...(extra.contactos || []),
                  ...this.contactoPersona(b?.id_responsable_area, 'Responsable'),
                  ...this.contactoPersona(b?.id_usuario_directo, 'Usuario directo')]
    };
  }

  // Desde datos del formulario o de una solicitud (llaves camelCase)
  formatoDesdeDatos(tipo: 'ALTA' | 'BAJA', d: any, extra: Partial<FormatoDoc> = {}): FormatoDoc {
    const ubic = (this.catalogs.ubicaciones || []).find((u: any) => String(u.id_ubicacion) === String(d?.idUbicacion));
    const marca = this.nombreDe(this.catalogs.marcas, 'id_marca', 'nombre_marca', d?.idMarca);
    const obs = [extra.observaciones, d?.caracteristicas ? 'Características: ' + d.caracteristicas : ''].filter(x => x).join('\n');
    return {
      tipo,
      folio: extra.folio || this.folio(tipo, null, d?.clave),
      sede: extra.sede || ubic?.nombre_sede || '',
      fecha: extra.fecha || this.fechaTexto(),
      filas: [this.filaDe(d?.clave, d?.articulo, marca, d?.modelo, d?.numeroSerie)],
      detalle: [
        { etiqueta: 'Clasificación', valor: this.nombreDe(this.catalogs.clasificaciones, 'id_clasificacion', 'nombre', d?.idClasificacion) },
        { etiqueta: 'Ubicación', valor: [ubic?.nombre_area, ubic?.nombre_sede].filter((x: any) => x).join(' · ') },
        { etiqueta: 'Responsable', valor: this.nombreDe(this.catalogs.personas, 'id_persona', 'nombre_completo', d?.idResponsableArea) },
        { etiqueta: 'Estado físico', valor: d?.estadoFisico || '' },
        { etiqueta: 'Piezas', valor: String(d?.cantidadPiezas || 1) },
        { etiqueta: 'Usuario directo', valor: this.nombreDe(this.catalogs.personas, 'id_persona', 'nombre_completo', d?.idUsuarioDirecto) }
      ],
      observaciones: obs,
      solicito: extra.solicito || '',
      autorizo: extra.autorizo || '',
      correoSugerido: extra.correoSugerido,
      contactos: [...(extra.contactos || []),
                  ...this.contactoPersona(d?.idResponsableArea, 'Responsable'),
                  ...this.contactoPersona(d?.idUsuarioDirecto, 'Usuario directo')]
    };
  }

  private contactoPersona(id: any, rol: string): { nombre: string; correo: string; rol: string }[] {
    if (id === null || id === undefined || id === '') return [];
    const p = (this.catalogs.personas || []).find((x: any) => String(x.id_persona) === String(id));
    return p?.correo ? [{ nombre: p.nombre_completo || p.correo, correo: p.correo, rol }] : [];
  }

  formatoDesdeSolicitud(s: any, bien: any): FormatoDoc {
    const tipo: 'ALTA' | 'BAJA' = s.tipoOperacion === 'BAJA' ? 'BAJA' : 'ALTA';
    const obs = [s.motivoSolicitud ? 'Motivo: ' + s.motivoSolicitud : '', s.motivoResolucion ? 'Comentario del administrador: ' + s.motivoResolucion : '']
      .filter(x => x).join('\n');
    const extra: Partial<FormatoDoc> = {
      folio: this.folio(tipo, s.idSolicitud),
      fecha: this.fechaTexto(s.fechaSolicitud),
      solicito: s.nombreSolicitante || s.correoSolicitante || '',
      autorizo: s.nombreRevisor || this.miNombre,
      sede: s.nombreSede || '',
      observaciones: obs,
      correoSugerido: s.correoSolicitante,
      contactos: s.correoSolicitante ? [{ nombre: s.nombreSolicitante || s.correoSolicitante, correo: s.correoSolicitante, rol: 'Solicitó' }] : []
    };
    if (tipo === 'ALTA') return this.formatoDesdeDatos('ALTA', s.datosPropuestos || {}, extra);
    if (bien) return { ...this.formatoDesdeBien('BAJA', bien, extra), sede: s.nombreSede || bien.sede || '', observaciones: obs };
    return {
      tipo, folio: extra.folio!, sede: s.nombreSede || '', fecha: extra.fecha!,
      filas: [this.filaDe(s.claveBien, s.articuloBien, '', '', '')],
      observaciones: obs, solicito: extra.solicito!, autorizo: extra.autorizo!, correoSugerido: s.correoSolicitante,
      contactos: extra.contactos
    };
  }

  // =========================
  // INVENTARIO ANUAL
  // =========================

  get verInventarioAnual(): boolean {
    return this.isAdmin || this.levantamientos.some(l => l.soyCapturista);
  }

  loadLevantamientos(): void {
    if (!this.logged) return;
    this.levCargando = true;
    this.http.get<any[]>(this.api + '/levantamientos').subscribe({
      next: rows => {
        this.levantamientos = rows || [];
        this.levCargando = false;
        if (this.levAbierto === null && this.levantamientos.length === 1) this.levAbierto = this.levantamientos[0].idLevantamiento;
      },
      error: () => { this.levCargando = false; }
    });
  }

  pctLev(l: any): number {
    return l.totalBienes ? Math.round((l.capturados / l.totalBienes) * 100) : 0;
  }

  abrirNuevoLevantamiento(): void {
    const ahora = new Date();
    const local = new Date(ahora.getTime() - ahora.getTimezoneOffset() * 60000).toISOString().slice(0, 16);
    this.candidatos = [];
    this.levError = '';
    this.levNuevo = {
      paso: 1, idSede: null, anio: ahora.getFullYear(),
      ciudad: '',
      direccion: '',
      fechaInicio: local, idUsuarioCapturista: null,
      participantes: [
        { nombre: 'C. Juan Carlos Zuno Muñiz', cargo: 'Jefe de Recursos Materiales' },
        { nombre: 'Mariana Aboytes Morales', cargo: 'Auxiliar administrativo' },
        { nombre: 'Gricelda Sanches Cruz', cargo: 'Jefa de departamento licitaciones' },
        { nombre: 'Ximena Cornejo Rivera', cargo: 'Auxiliar administrativo' }
      ],
      pregunta: { texto: '¿En qué estado está el bien?', opciones: ['Bueno', 'Regular', 'Malo', 'No funciona'] },
      nuevaOpcion: ''
    };
  }

  // Al elegir la sede, llena la dirección con la registrada en el catálogo de Sedes (se puede editar).
  ponerDireccionSede(): void {
    const n = this.levNuevo;
    const s = (this.catalogs.sedes || []).find((x: any) => x.id_sede === n?.idSede);
    if (!n || !s) return;
    n.direccion = s.direccion_completa || '';
    // Ciudad sugerida "Sede, Entidad" (solo si el usuario no la escribió a mano)
    if (!String(n.ciudad || '').trim() || n.ciudad === n._ciudadAuto) {
      const t = (v: string) => String(v || '').toLowerCase().replace(/(^|\s)\S/g, c => c.toUpperCase());
      n.ciudad = n._ciudadAuto = [t(s.nombre_sede), s.nombre_entidad && s.nombre_entidad !== s.nombre_sede ? t(s.nombre_entidad) : '']
        .filter(x => !!x).join(', ');
    }
  }

  cargarCandidatos(): void {
    this.candidatos = [];
    if (!this.levNuevo?.idSede) return;
    this.levNuevo.idUsuarioCapturista = null;
    this.http.get<any[]>(this.api + '/levantamientos/candidatos?idSede=' + this.levNuevo.idSede).subscribe({
      next: rows => this.candidatos = rows || [],
      error: e => this.levError = this.mensajeError(e, 'No se pudieron cargar las personas de la sede.')
    });
  }

  siguientePasoLev(): void {
    const n = this.levNuevo;
    this.levError = '';
    if (n.paso === 1 && !n.idSede) { this.levError = 'Selecciona la sede.'; return; }
    if (n.paso === 2 && (!String(n.ciudad || '').trim() || !n.fechaInicio)) { this.levError = 'Escribe la ciudad y la fecha de inicio.'; return; }
    if (n.paso === 3) {
      if (!n.idUsuarioCapturista) { this.levError = 'Elige quién va a capturar.'; return; }
      if (!n.participantes.some((p: any) => String(p.nombre || '').trim())) { this.levError = 'Agrega al menos un participante.'; return; }
    }
    n.paso++;
  }

  agregarOpcion(): void {
    const v = String(this.levNuevo.nuevaOpcion || '').trim();
    if (v && !this.levNuevo.pregunta.opciones.includes(v)) this.levNuevo.pregunta.opciones.push(v);
    this.levNuevo.nuevaOpcion = '';
  }

  crearLevantamiento(): void {
    const n = this.levNuevo;
    this.levGuardando = true;
    this.levError = '';
    this.http.post<any>(this.api + '/levantamientos', {
      idSede: n.idSede, anio: n.anio, ciudad: n.ciudad, direccion: n.direccion, fechaInicio: n.fechaInicio,
      idUsuarioCapturista: n.idUsuarioCapturista,
      participantes: n.participantes.filter((p: any) => String(p.nombre || '').trim()),
      pregunta: n.pregunta
    }).subscribe({
      next: r => {
        this.levGuardando = false;
        this.levNuevo = null;
        this.levAbierto = r.idLevantamiento;
        this.loadLevantamientos();
        this.notificar('exito', 'Levantamiento iniciado', 'Inventario ' + r.anio + ' · ' + r.sede + '. El acta de inicio quedó guardada en Documentos.', 3500,
          { texto: '📄 Ver acta de inicio', fn: () => this.abrirDocumentoGuardado(r.idDocumentoInicio) });
      },
      error: e => { this.levGuardando = false; this.levError = this.mensajeError(e, 'No fue posible crear el levantamiento.'); }
    });
  }

  // Pendientes: sin revisar ni marcar (cerrado: todo lo no revisado queda "no localizado")
  pendientesLev(l: any): number {
    const t = l.totalBienes || 0, rev = l.capturados || 0;
    return Math.max(0, l.estado === 'CERRADO' ? t - rev : t - rev - (l.noEncontrados || 0));
  }

  // Lleva a Documentos > Resguardos, filtrado por este levantamiento
  verResguardosLev(l: any): void {
    this.section = 'documentos';
    this.docsVista = 'RESGUARDO';
    this.docsBusqueda = 'Inventario ' + l.anio + ' ' + l.sede;
    this.loadDocumentos();
    window.scrollTo({ top: 0 });
  }

  // Abre el resguardo (borrador o vigente) de una persona en un levantamiento
  abrirResguardo(lev: any, idPersona: any): void {
    this.http.get<any[]>(this.api + '/documentos?tipo=RESGUARDO').subscribe({
      next: rows => {
        const d = (rows || []).find(x => x.idLevantamiento === lev.idLevantamiento && x.idPersona === idPersona);
        if (d) this.abrirDocumentoGuardado(d.idDocumento);
        else this.notificar('info', 'Sin resguardo', 'Esta persona todavía no tiene bienes revisados en este levantamiento.');
      },
      error: e => this.notificar('error', 'No se abrió el resguardo', this.mensajeError(e, 'Inténtalo de nuevo.'), 0)
    });
  }

  cerrarLevantamiento(l: any, yaRevisoPendientes = false): void {
    const faltan = this.pendientesLev(l);
    const marcados = l.noEncontrados || 0;
    if (faltan > 0 && !yaRevisoPendientes) {
      this.abrirAvance(l, { filtro: 'pendientes', antesDeCerrar: true });
      return;
    }
    this.abrirDialogo({
      tipo: 'peligro', icono: '✓',
      titulo: '¿Cerrar el levantamiento?',
      mensaje: 'Se revisaron ' + l.capturados + ' de ' + l.totalBienes + ' bienes.' +
        (faltan + marcados > 0 ? ' Los ' + (faltan + marcados) + ' restantes (' + marcados + ' marcados como no encontrados) quedarán como "no localizados" en el acta.' : '') +
        ' Ya no se podrá capturar; se emitirán el Acta de cierre y el resguardo de cada persona.',
      textoConfirmar: 'Sí, cerrar y generar acta',
      accion: terminar => {
        this.http.post<any>(this.api + '/levantamientos/' + l.idLevantamiento + '/cerrar', {}).subscribe({
          next: r => {
            terminar();
            this.loadLevantamientos();
            this.notificar('exito', 'Levantamiento cerrado', 'El acta de cierre y los resguardos quedaron emitidos en Documentos.', 3500,
              { texto: '📄 Ver acta de cierre', fn: () => this.abrirDocumentoGuardado(r.idDocumentoCierre) });
          },
          error: e => terminar(this.mensajeError(e, 'No fue posible cerrar el levantamiento.'))
        });
      }
    });
  }

  eliminarLevantamiento(l: any): void {
    const fotos = l.fotosNuevas || 0;
    this.abrirDialogo({
      tipo: 'peligro', icono: '🗑',
      titulo: '¿Eliminar este inventario?',
      mensaje: 'Se borrará el inventario de ' + (l.sede || 'la sede') + ' ' + (l.anio || '') +
        ' con sus ' + (l.capturados || 0) + ' bienes revisados, sus actas' +
        (fotos > 0 ? ' y las ' + fotos + ' fotos que se tomaron en él (también de Cloudinary). Esos bienes volverán a pedir foto.' : '.') +
        ' Esta acción no se puede deshacer.',
      textoConfirmar: 'Sí, eliminar todo',
      accion: terminar => {
        this.http.delete<any>(this.api + '/levantamientos/' + l.idLevantamiento).subscribe({
          next: r => {
            terminar();
            if (this.levAbierto === l.idLevantamiento) this.levAbierto = null;
            // Los bienes que perdieron su foto deben volver a pedirla en Inventario
            if ((r?.fotosBorradas || 0) + (r?.fotosNoBorradas || 0) > 0) this.loadBienes();
            this.loadLevantamientos();
            this.notificar('exito', 'Inventario eliminado',
              'Se borraron ' + (r?.registros || 0) + ' registros y ' + (r?.fotosBorradas || 0) + ' fotos.' +
              (r?.fotosNoBorradas ? ' ' + r.fotosNoBorradas + ' fotos no se pudieron borrar de Cloudinary; puedes quitarlas desde su página.' : ''));
          },
          error: e => terminar(this.mensajeError(e, 'No fue posible eliminar el inventario.'))
        });
      }
    });
  }

  // ----- Avance (monitoreo) -----

  abrirAvance(l: any, opciones: { filtro?: string; antesDeCerrar?: boolean } = {}): void {
    this.detenerTimerAvance();
    this.anualVista = 'avance';
    this.av = this.nuevoAvance(l, false, opciones.filtro);
    this.av.antesDeCerrar = !!opciones.antesDeCerrar;
    this.av.cargando = true;
    this.avCargar();
    if (l.estado === 'EN_CURSO') {
      // Para los observadores: se refresca solo cada 30 s mientras la vista siga abierta
      this.avTimer = setInterval(() => {
        if (!this.av || this.av.enCaptura || this.section !== 'anual') { this.detenerTimerAvance(); return; }
        this.avCargar();
      }, 30000);
    }
    window.scrollTo({ top: 0 });
  }

  abrirAvanceCaptura(filtro = 'todos'): void {
    if (!this.cap) return;
    this.detenerEscaner();
    this.av = this.nuevoAvance(this.cap.lev, true, filtro);
    this.av.bienes = this.cap.bienes;
    if (this.cap.grupo) {
      this.av.agrupar = this.cap.grupo.modo;
      if (filtro === 'pendientes') this.av.areaAbierta = this.cap.grupo.key;
    }
    this.avRecalcular();
    window.scrollTo({ top: 0 });
  }

  private nuevoAvance(l: any, enCaptura: boolean, filtro = 'todos'): any {
    return {
      lev: l, enCaptura, bienes: [], cargando: false, error: '', filtro, busqueda: '', agrupar: 'persona',
      areaAbierta: null, limites: {}, detalle: null, antesDeCerrar: false, actualizado: null,
      vista: { total: 0, revisados: 0, pct: 0, cuentas: {}, areas: [], resultados: [] }
    };
  }

  private avCargar(): void {
    const av = this.av;
    this.http.get<any[]>(this.api + '/levantamientos/' + av.lev.idLevantamiento + '/bienes').subscribe({
      next: rows => {
        if (this.av !== av) return;
        av.bienes = rows || [];
        av.cargando = false;
        av.error = '';
        av.actualizado = new Date();
        this.avRecalcular();
      },
      error: e => {
        if (this.av !== av) return;
        av.cargando = false;
        av.error = this.mensajeError(e, 'No se pudo cargar el avance.');
      }
    });
  }

  cerrarAvance(): void {
    this.detenerTimerAvance();
    const enCaptura = this.av?.enCaptura;
    this.av = null;
    if (!enCaptura) {
      this.anualVista = 'lista';
      this.loadLevantamientos();
    }
    window.scrollTo({ top: 0 });
  }

  cerrarDesdeAvance(): void {
    const l = this.av.lev;
    l.capturados = this.av.vista.revisados;
    l.noEncontrados = this.av.vista.cuentas.noenc || 0;
    this.cerrarAvance();
    this.cerrarLevantamiento(l, true);
  }

  private detenerTimerAvance(): void {
    if (this.avTimer) { clearInterval(this.avTimer); this.avTimer = null; }
  }

  avEstado(b: any): 'ok' | 'fuera' | 'pendiente' | 'no' {
    if (b.estatus === 'NO_ENCONTRADO') return 'no';
    if (b.estatus !== 'REVISADO') return 'pendiente';
    return b.fuera_de_lugar ? 'fuera' : 'ok';
  }

  private avCoincide(b: any, filtro: string): boolean {
    switch (filtro) {
      case 'pendientes': return !b.estatus;
      case 'revisados': return b.estatus === 'REVISADO';
      case 'noenc': return b.estatus === 'NO_ENCONTRADO';
      case 'fuera': return b.estatus === 'REVISADO' && !!b.fuera_de_lugar;
      case 'sinfoto': return !b.tiene_foto;
      case 'sinetiq': return !b.etiqueta;
      default: return true;
    }
  }

  avAgrupar(modo: 'persona' | 'area'): void {
    this.av.agrupar = modo;
    this.av.areaAbierta = null;
    this.av.limites = {};
    this.avRecalcular();
  }

  avFiltrar(filtro: string): void {
    this.av.filtro = filtro;
    this.av.limites = {};
    this.avRecalcular();
  }

  avLimite(key: any): number {
    return this.av?.limites[key] || 40;
  }

  trackArea(_: number, a: any): any {
    return a.key;
  }

  // Agrupa por persona (resguardo) o por área registrada; se calcula solo cuando cambian datos o filtros
  avRecalcular(): void {
    const av = this.av;
    if (!av) return;
    const porPersona = av.agrupar === 'persona';
    const cuentas: any = { todos: 0, pendientes: 0, revisados: 0, noenc: 0, fuera: 0, sinfoto: 0, sinetiq: 0 };
    const mapa = new Map<any, any>();
    for (const b of av.bienes) {
      for (const f of Object.keys(cuentas)) if (this.avCoincide(b, f)) cuentas[f]++;
      const key = (porPersona ? b.id_persona : b.id_ubicacion) ?? 'sin';
      let a = mapa.get(key);
      if (!a) {
        a = { key, nombre: porPersona ? (b.persona || 'Sin resguardante') : (b.ubicacion || 'Sin ubicación'), total: 0, revisados: 0, no: 0, fuera: 0, items: [] };
        mapa.set(key, a);
      }
      a.total++;
      if (b.estatus === 'REVISADO') a.revisados++;
      if (b.estatus === 'NO_ENCONTRADO') a.no++;
      if (b.estatus === 'REVISADO' && b.fuera_de_lugar) a.fuera++;
      if (this.avCoincide(b, av.filtro)) a.items.push(b);
    }
    const orden: any = { pendiente: 0, no: 1, fuera: 2, ok: 3 };
    const areas = Array.from(mapa.values())
      .filter(a => av.filtro === 'todos' || a.items.length)
      .sort((x, y) => Number(x.key === 'sin') - Number(y.key === 'sin') || x.nombre.localeCompare(y.nombre, 'es'));
    for (const a of areas) {
      a.pct = a.total ? Math.round((a.revisados / a.total) * 100) : 0;
      a.items.sort((x: any, y: any) => orden[this.avEstado(x)] - orden[this.avEstado(y)]
        || String(x.clave || '').localeCompare(String(y.clave || ''), 'es', { numeric: true }));
    }
    const q = String(av.busqueda || '').toLowerCase().trim();
    const resultados = q
      ? av.bienes.filter((b: any) => this.avCoincide(b, av.filtro)
          && [b.clave, b.articulo, b.numero_serie, b.marca, b.modelo, b.persona, b.etiqueta].some(x => String(x || '').toLowerCase().includes(q))).slice(0, 60)
      : [];
    const total = av.bienes.length, revisados = cuentas.revisados;
    av.vista = { total, revisados, pct: total ? Math.round((revisados / total) * 100) : 0, cuentas, areas, resultados };
  }

  avTocarBien(b: any): void {
    if (this.av.enCaptura) {
      if (this.cap?.lev?.estado !== 'EN_CURSO') return;
      if (this.cap.etiquetaPorAsignar) {
        if (b.etiqueta) { this.notificar('info', 'Ese bien ya tiene etiqueta', b.clave + ' ya tiene la ' + b.etiqueta + '. Elige otro.'); return; }
        this.asignarEtiquetaCap(this.cap.etiquetaPorAsignar, b);
        return;
      }
      this.av = null;
      this.abrirBienCaptura(b);
      this.cap.desdeLista = true;
      return;
    }
    this.av.detalle = this.av.detalle === b.id_bien ? null : b.id_bien;
  }

  // Hoja para recorrer las areas con los bienes que faltan (no es el acta)
  avPdfPendientes(): void {
    const av = this.av;
    const porPersona = av.agrupar === 'persona';
    const grupos = new Map<string, any[]>();
    for (const b of av.bienes) {
      if (b.estatus) continue;
      const k = porPersona ? (b.persona || 'Sin resguardante') : (b.ubicacion || 'Sin ubicación');
      if (!grupos.has(k)) grupos.set(k, []);
      grupos.get(k)!.push(b);
    }
    const areas = Array.from(grupos.keys()).sort((x, y) => x.localeCompare(y, 'es'));
    const total = areas.reduce((n, k) => n + grupos.get(k)!.length, 0);
    const doc = new jsPDF({ unit: 'mm', format: 'letter', compress: true });
    const W = 215.9, M = 14, ABAJO = 262;
    const col = { chk: M, clave: M + 7, art: M + 33, marca: M + 103, serie: M + 150 };
    let y = 0;
    const encabezado = () => {
      doc.setFont('helvetica', 'bold'); doc.setFontSize(14); doc.setTextColor(20, 36, 58);
      doc.text('Lista de pendientes por ' + (porPersona ? 'persona' : 'área'), M, 18);
      doc.setFont('helvetica', 'normal'); doc.setFontSize(9.5); doc.setTextColor(95, 110, 130);
      doc.text('Inventario ' + av.lev.anio + ' · ' + av.lev.sede + ' · ' + total + ' bienes pendientes · generado ' +
        new Date().toLocaleString('es-MX', { dateStyle: 'short', timeStyle: 'short' }), M, 24);
      doc.setDrawColor(45, 93, 147); doc.setLineWidth(0.6); doc.line(M, 27, W - M, 27);
      y = 34;
    };
    const titulos = () => {
      doc.setFont('helvetica', 'bold'); doc.setFontSize(8); doc.setTextColor(100, 130, 166);
      doc.text('CLAVE', col.clave, y); doc.text('ARTÍCULO', col.art, y); doc.text(porPersona ? 'ÁREA' : 'MARCA / MODELO', col.marca, y); doc.text('SERIE', col.serie, y);
      y += 4;
    };
    encabezado();
    for (const area of areas) {
      const lista = grupos.get(area)!.slice().sort((x, z) => String(x.clave || '').localeCompare(String(z.clave || ''), 'es', { numeric: true }));
      if (y > ABAJO - 20) { doc.addPage(); encabezado(); }
      doc.setFillColor(238, 243, 249); doc.rect(M, y - 4.5, W - 2 * M, 7, 'F');
      doc.setFont('helvetica', 'bold'); doc.setFontSize(10.5); doc.setTextColor(20, 36, 58);
      doc.text(this.ajustarTexto(doc, area, 150), M + 2, y);
      doc.text(lista.length + (lista.length === 1 ? ' pendiente' : ' pendientes'), W - M - 2, y, { align: 'right' });
      y += 7;
      titulos();
      for (const b of lista) {
        if (y > ABAJO) { doc.addPage(); encabezado(); doc.setFont('helvetica', 'bold'); doc.setFontSize(9.5); doc.setTextColor(20, 36, 58); doc.text(area + ' (continúa)', M, y); y += 6; titulos(); }
        doc.setDrawColor(150, 160, 175); doc.setLineWidth(0.3); doc.rect(col.chk, y - 3, 3.6, 3.6);
        doc.setFont('helvetica', 'bold'); doc.setFontSize(9); doc.setTextColor(20, 36, 58);
        doc.text(this.ajustarTexto(doc, b.clave, 24), col.clave, y);
        doc.setFont('helvetica', 'normal'); doc.setTextColor(45, 55, 70);
        doc.text(this.ajustarTexto(doc, b.articulo, 67), col.art, y);
        doc.text(this.ajustarTexto(doc, (porPersona ? b.ubicacion : [b.marca, b.modelo].filter(x => this.esTexto(x)).join(' / ')) || '—', 44), col.marca, y);
        doc.text(this.ajustarTexto(doc, b.numero_serie || '—', W - M - col.serie), col.serie, y);
        doc.setDrawColor(232, 236, 241); doc.line(M, y + 1.8, W - M, y + 1.8);
        y += 6;
      }
      y += 4;
    }
    const paginas = doc.getNumberOfPages();
    for (let i = 1; i <= paginas; i++) {
      doc.setPage(i);
      doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor(140, 150, 165);
      doc.text('DOZZER · Página ' + i + ' de ' + paginas, W / 2, 272, { align: 'center' });
    }
    const nombre = ('pendientes-' + av.lev.sede + '-' + av.lev.anio).toLowerCase().replace(/[^a-z0-9-]+/g, '-');
    doc.save(nombre + '.pdf');
  }

  // ----- Captura -----

  get camaraEnVivo(): boolean {
    return !!(window.isSecureContext && navigator.mediaDevices && navigator.mediaDevices.getUserMedia);
  }

  iniciarCaptura(l: any): void {
    this.anualVista = 'captura';
    this.cap = { lev: l, bienes: [], cargando: true, paso: 'area', area: null, busquedaArea: '', bien: null,
                 foto: null, fotoUrl: '', respuesta: '', guardando: false, error: '', ultimo: '', clave: '',
                 verClave: false, verLista: false, escaneando: false,
                 modo: 'persona', grupo: null, agregado: null, ofrecerNuevo: false, claveSugerida: '', desdeLista: false, cambios: false };
    this.capOpciones = { areas: [], personas: [], clasificaciones: [], marcas: [] };
    this.http.get<any>(this.api + '/levantamientos/' + l.idLevantamiento + '/opciones').subscribe({
      next: o => this.capOpciones = { areas: [], personas: [], clasificaciones: [], marcas: [], ...(o || {}) },
      error: () => { /* sin opciones: el formulario de bien nuevo avisará */ }
    });
    this.http.get<any[]>(this.api + '/levantamientos/' + l.idLevantamiento + '/bienes').subscribe({
      next: rows => { this.cap.bienes = rows || []; this.cap.cargando = false; },
      error: e => { this.cap.cargando = false; this.cap.error = this.mensajeError(e, 'No se pudieron cargar los bienes.'); }
    });
    window.scrollTo({ top: 0 });
  }

  salirCaptura(): void {
    this.detenerEscaner();
    this.limpiarFotoCaptura();
    // Hubo bienes nuevos o traídos: la lista de Bienes debe reflejarlos
    if (this.cap?.cambios) this.loadBienes();
    this.cap = null;
    this.anualVista = 'lista';
    this.loadLevantamientos();
  }

  get capRevisados(): number {
    return this.cap ? this.cap.bienes.filter((b: any) => b.estatus === 'REVISADO').length : 0;
  }

  get capPct(): number {
    return this.cap?.bienes.length ? Math.round((this.capRevisados / this.cap.bienes.length) * 100) : 0;
  }

  // Personas o áreas de la sede con su avance (hechos = revisados + no encontrados)
  get gruposCaptura(): any[] {
    if (!this.cap) return [];
    const porPersona = this.cap.modo === 'persona';
    const mapa = new Map<any, any>();
    for (const b of this.cap.bienes) {
      const id = (porPersona ? b.id_persona : b.id_ubicacion) ?? null;
      const key = id ?? 'sin';
      const g = mapa.get(key) || { key, id, nombre: porPersona ? (b.persona || 'Sin resguardante') : (b.ubicacion || 'Sin ubicación'), total: 0, hechos: 0 };
      g.total++;
      if (b.estatus) g.hechos++;
      mapa.set(key, g);
    }
    const q = String(this.cap.busquedaArea || '').toLowerCase().trim();
    return Array.from(mapa.values())
      .filter(g => !q || g.nombre.toLowerCase().includes(q))
      .sort((a, b) => Number(a.key === 'sin') - Number(b.key === 'sin') || a.nombre.localeCompare(b.nombre, 'es'));
  }

  // g = persona o área elegida; null = toda la sede
  elegirGrupo(g: any): void {
    this.cap.grupo = g ? { ...g, modo: this.cap.modo } : null;
    this.cap.area = g && this.cap.modo === 'area' ? { id: g.id, nombre: g.nombre } : null;
    this.cap.paso = 'escanear';
    this.cap.error = '';
  }

  private enGrupo(b: any): boolean {
    const g = this.cap?.grupo;
    if (!g) return true;
    return ((g.modo === 'persona' ? b.id_persona : b.id_ubicacion) ?? 'sin') === g.key;
  }

  get bienesGrupo(): any[] {
    return this.cap ? this.cap.bienes.filter((b: any) => this.enGrupo(b)) : [];
  }

  get pendientesGrupo(): any[] {
    return this.bienesGrupo.filter((b: any) => !b.estatus);
  }

  get grupoActualNombre(): string {
    if (this.cap?.grupo) return this.cap.grupo.nombre;
    return this.cap?.modo === 'area' ? 'Todas las áreas' : 'Todas las personas';
  }

  esTexto(x: any): boolean {
    return !!(x && String(x).trim());
  }

  // Identifica el bien con lo que traiga el QR (enlace de la ficha, texto o la clave)
  private identificarBien(texto: string): any {
    const { id, clave, etiqueta } = this.datosQr(texto);
    const lista = this.cap.bienes;
    if (etiqueta) return lista.find((b: any) => b.etiqueta === etiqueta) || null;
    return (id && lista.find((b: any) => String(b.codigo_qr || '').replace(/-/g, '').substring(0, 8).toUpperCase() === id))
      || (clave && lista.find((b: any) => String(b.clave || '').toUpperCase() === clave.toUpperCase()))
      || null;
  }

  private datosQr(texto: string): { id: string; clave: string; etiqueta: string } {
    const t = String(texto || '').trim();
    let id = '', clave = '';
    const etiqueta = this.codigoDeTexto(t);
    if (etiqueta) return { id: '', clave: '', etiqueta };
    const hash = t.includes('#') ? t.split('#')[1] : t.includes('?') ? t.split('?')[1] : '';
    if (hash) {
      try { const ps = new URLSearchParams(hash); id = (ps.get('i') || '').toUpperCase(); clave = ps.get('c') || ''; } catch { /* no es enlace */ }
    }
    if (!clave) {
      const m = t.match(/Clave:\s*(.+)/i);
      if (m) clave = m[1].trim();
    }
    if (!id && !clave) clave = t;
    return { id, clave, etiqueta: '' };
  }

  private alLeerQr(texto: string): void {
    const b = this.identificarBien(texto);
    if (b) { this.abrirBienCaptura(b); return; }
    const { etiqueta } = this.datosQr(texto);
    if (etiqueta) { this.etiquetaEscaneada(etiqueta); return; }
    this.buscarFueraDeSede(texto);
  }

  // Etiqueta que no está en la lista de la sede: libre (se asigna), de otra sede (se ofrece traer) o inexistente
  private etiquetaEscaneada(codigo: string): void {
    const c = this.cap;
    c.error = ''; c.ultimo = ''; c.ofrecerNuevo = false;
    this.http.get<any>(this.api + '/etiquetas/' + encodeURIComponent(codigo)).subscribe({
      next: r => {
        if (r.estado === 'LIBRE') {
          // Bien recién agregado sin etiqueta: se le asigna directo
          const ag = c.agregado?.bien;
          if (ag && !ag.etiqueta) { this.asignarEtiquetaCap(codigo, ag); return; }
          c.etiquetaPorAsignar = codigo;
          this.abrirAvanceCaptura('sinetiq');
          return;
        }
        if (r.estado === 'ASIGNADA' && r.clave) { this.abrirBuscarGeneral(r.clave); return; }
        c.error = r.estado === 'ANULADA' ? 'La etiqueta ' + codigo + ' está anulada. Usa otra.' : 'La etiqueta ' + codigo + ' no existe en DOZZER.';
      },
      error: e => c.error = this.mensajeError(e, 'No se pudo consultar la etiqueta.')
    });
  }

  // Liga la etiqueta escaneada al bien y abre su captura
  asignarEtiquetaCap(codigo: string, b: any): void {
    const c = this.cap;
    this.http.post<any>(this.api + '/etiquetas/' + encodeURIComponent(codigo) + '/asignar', { idBien: b.id_bien }).subscribe({
      next: () => {
        b.etiqueta = codigo;
        c.etiquetaPorAsignar = null;
        this.av = null;
        this.abrirBienCaptura(b);
        c.ultimo = '';
        this.notificar('exito', 'Etiqueta asignada', codigo + ' → ' + b.clave + ' · ' + b.articulo, 2200);
      },
      error: e => { c.etiquetaPorAsignar = null; c.error = this.mensajeError(e, 'No se pudo asignar la etiqueta.'); this.av = null; }
    });
  }

  get sinEtiquetaGrupo(): number {
    return this.bienesGrupo.filter((b: any) => !b.etiqueta).length;
  }

  // No está en esta sede: se busca en todas; si aparece en otra se ofrece traerlo, si no, registrarlo como nuevo
  private buscarFueraDeSede(texto: string): void {
    const c = this.cap;
    const { id, clave } = this.datosQr(texto);
    const q = (id || clave || '').trim();
    c.error = '';
    c.ultimo = '';
    c.ofrecerNuevo = false;
    if (q.length < 2) { c.error = 'No se reconoció el código.'; return; }
    this.http.get<any[]>(this.api + '/levantamientos/' + c.lev.idLevantamiento + '/buscar', { params: { q } }).subscribe({
      next: rows => {
        if ((rows || []).some(r => !r.en_esta_sede)) { this.abrirBuscarGeneral(q, rows || []); return; }
        c.error = '"' + (clave || q) + '" no está registrado en ninguna sede.';
        c.ofrecerNuevo = true;
        c.claveSugerida = !id && clave ? clave : '';
      },
      error: e => c.error = this.mensajeError(e, 'No se pudo buscar el bien.')
    });
  }

  abrirBienCaptura(b: any): void {
    this.detenerEscaner();
    this.limpiarFotoCaptura();
    this.cap.bien = b;
    this.cap.respuesta = '';
    this.cap.error = '';
    this.cap.ultimo = '';
    this.cap.agregado = null;
    this.cap.ofrecerNuevo = false;
    this.cap.desdeLista = false;
    this.cap.paso = 'bien';
    window.scrollTo({ top: 0 });
  }

  cancelarBienCaptura(): void {
    this.limpiarFotoCaptura();
    this.cap.bien = null;
    this.cap.paso = 'escanear';
  }

  buscarPorClave(): void {
    const b = this.identificarBien(this.cap.clave);
    if (!b) { this.buscarFueraDeSede(this.cap.clave); return; }
    this.cap.clave = '';
    this.abrirBienCaptura(b);
  }

  async iniciarEscaner(): Promise<void> {
    this.cap.error = '';
    try {
      this.camStream = await navigator.mediaDevices.getUserMedia({
        video: { facingMode: { ideal: 'environment' }, width: { ideal: 1280 }, height: { ideal: 720 } }, audio: false
      });
    } catch (e: any) {
      this.cap.error = e?.name === 'NotAllowedError'
        ? 'DOZZER no tiene permiso para usar la cámara. Actívalo en los ajustes del celular (Ajustes › DOZZER o Safari › Cámara) o usa "Leer QR desde una foto".'
        : 'No se pudo abrir la cámara. Usa "Leer QR desde una foto".';
      return;
    }
    this.cap.escaneando = true;
    // Espera a que el cuadro de video exista en pantalla (aparece al activar el escáner)
    let video: HTMLVideoElement | undefined;
    for (let i = 0; i < 30 && !video; i++) {
      await new Promise(r => setTimeout(r, 50));
      video = this.camQr?.nativeElement;
    }
    if (!video || !this.camStream) { this.detenerEscaner(); this.cap.error = 'No se pudo mostrar la cámara. Usa "Leer QR desde una foto".'; return; }
    // iPhone: el video solo se reproduce solo si está silenciado y en línea (estas propiedades deben ir antes del stream)
    video.muted = true;
    video.setAttribute('muted', '');
    video.setAttribute('playsinline', '');
    video.setAttribute('autoplay', '');
    video.srcObject = this.camStream;
    try { await video.play(); } catch { /* algunos iPhone arrancan con el autoplay */ }
    const inicio = Date.now();
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d', { willReadFrequently: true } as any) as CanvasRenderingContext2D;
    const v = video;
    const leer = () => {
      if (!this.cap?.escaneando) return;
      if (v.readyState >= v.HAVE_CURRENT_DATA && v.videoWidth) {
        const escala = Math.min(1, 640 / v.videoWidth);
        canvas.width = Math.round(v.videoWidth * escala);
        canvas.height = Math.round(v.videoHeight * escala);
        ctx.drawImage(v, 0, 0, canvas.width, canvas.height);
        const img = ctx.getImageData(0, 0, canvas.width, canvas.height);
        const qr = jsQR(img.data, img.width, img.height, { inversionAttempts: 'dontInvert' });
        if (qr?.data) { this.alLeerQr(qr.data); return; }
      } else if (Date.now() - inicio > 6000) {
        // La cámara nunca mandó imagen (se queda en negro): se cierra y se ofrece la otra opción
        this.detenerEscaner();
        this.cap.error = 'La cámara no respondió. Ciérrala y vuelve a intentar, o usa "Leer QR desde una foto".';
        return;
      } else if (v.paused) {
        v.play().catch(() => {});
      }
      this.camTimer = setTimeout(leer, 180);
    };
    leer();
  }

  detenerEscaner(): void {
    clearTimeout(this.camTimer);
    if (this.camStream) this.camStream.getTracks().forEach(t => t.stop());
    this.camStream = null;
    if (this.cap) this.cap.escaneando = false;
  }

  async leerQrDeFoto(ev: Event): Promise<void> {
    const input = ev.target as HTMLInputElement;
    const archivo = input.files?.[0];
    input.value = '';
    if (!archivo) return;
    this.cap.error = '';
    try {
      const bmp = await this.cargarImagen(archivo);
      const escala = Math.min(1, 1400 / Math.max(bmp.width, bmp.height));
      const canvas = document.createElement('canvas');
      canvas.width = Math.round(bmp.width * escala);
      canvas.height = Math.round(bmp.height * escala);
      const ctx = canvas.getContext('2d') as CanvasRenderingContext2D;
      ctx.drawImage(bmp, 0, 0, canvas.width, canvas.height);
      const img = ctx.getImageData(0, 0, canvas.width, canvas.height);
      const qr = jsQR(img.data, img.width, img.height, { inversionAttempts: 'attemptBoth' });
      if (qr?.data) this.alLeerQr(qr.data);
      else this.cap.error = 'No se encontró un QR en la foto. Acércate más y que salga enfocado.';
    } catch {
      this.cap.error = 'No se pudo leer la foto.';
    }
  }

  private async cargarImagen(archivo: Blob): Promise<ImageBitmap | HTMLImageElement> {
    try {
      return await createImageBitmap(archivo, { imageOrientation: 'from-image' } as any);
    } catch {
      return await new Promise((ok, mal) => {
        const img = new Image();
        img.onload = () => ok(img);
        img.onerror = mal;
        img.src = URL.createObjectURL(archivo);
      });
    }
  }

  // Foto del bien: se reduce a 1280 px y se comprime (~150-300 KB) antes de subirla
  async tomarFotoBien(ev: Event): Promise<void> {
    const input = ev.target as HTMLInputElement;
    const archivo = input.files?.[0];
    input.value = '';
    if (!archivo) return;
    try {
      const img = await this.cargarImagen(archivo);
      const escala = Math.min(1, 1280 / Math.max(img.width, img.height));
      const canvas = document.createElement('canvas');
      canvas.width = Math.round(img.width * escala);
      canvas.height = Math.round(img.height * escala);
      (canvas.getContext('2d') as CanvasRenderingContext2D).drawImage(img, 0, 0, canvas.width, canvas.height);
      const blob: Blob = await new Promise((ok, mal) => canvas.toBlob(b => b ? ok(b) : mal(), 'image/jpeg', 0.72));
      this.limpiarFotoCaptura();
      this.cap.foto = blob;
      this.cap.fotoUrl = URL.createObjectURL(blob);
      this.cap.error = '';
    } catch {
      this.cap.error = 'No se pudo procesar la foto. Inténtalo de nuevo.';
    }
  }

  private limpiarFotoCaptura(): void {
    if (this.cap?.fotoUrl) URL.revokeObjectURL(this.cap.fotoUrl);
    if (this.cap) { this.cap.foto = null; this.cap.fotoUrl = ''; }
  }

  // La lista de bienes se descarga una vez; la foto pudo tomarse después (Inventario anual), así que se consulta al abrir.
  private refrescarFotoBien(b: any): void {
    if (!b?.id_bien) return;
    this.http.get<any>(this.api + '/bienes/' + b.id_bien).subscribe({
      next: f => {
        if (!f) return;
        b.tiene_foto = f.tiene_foto; b.foto_mini = f.foto_mini; b.foto_url = f.foto_url;
      },
      error: () => { /* sin conexión: se queda con lo que ya tenía */ }
    });
  }

  // Al tomar una foto en el Inventario anual, actualiza también el bien en la lista de Inventario.
  private actualizarFotoEnInventario(idBien: number, mini: string, url: string): void {
    const enLista = this.bienes.find(x => x.id_bien === idBien);
    if (enLista) { enLista.tiene_foto = true; enLista.foto_mini = mini; enLista.foto_url = url; }
  }

  ampliarFoto(b: any): void {
    const url = b?.foto_url || b?.foto_mini;
    if (url) this.fotoAmpliada = { url, titulo: [b.clave, b.articulo].filter((x: any) => !!x).join(' · ') };
  }

  guardarCaptura(): void {
    const c = this.cap, b = c.bien;
    if (!b || c.guardando) return;
    const fd = new FormData();
    fd.append('idBien', String(b.id_bien));
    fd.append('respuesta', c.respuesta);
    if (c.area?.id) fd.append('idUbicacion', String(c.area.id));
    if (!b.tiene_foto && c.foto) fd.append('foto', c.foto, 'foto.jpg');
    c.guardando = true;
    c.error = '';
    this.http.post<any>(this.api + '/levantamientos/' + c.lev.idLevantamiento + '/registros', fd).subscribe({
      next: r => {
        c.guardando = false;
        b.respuesta = c.respuesta;
        b.estatus = 'REVISADO';
        b.fuera_de_lugar = !!r?.fueraDeLugar;
        b.ubicacion_encontrada = r?.fueraDeLugar ? (c.area?.nombre || null) : null;
        b.fecha_registro = new Date().toISOString();
        if (r?.fotoNueva) {
          b.tiene_foto = true; b.foto_mini = r.foto_mini; b.foto_url = r.foto_url;
          this.actualizarFotoEnInventario(b.id_bien, r.foto_mini, r.foto_url);
        }
        c.ultimo = b.clave + ' · ' + c.respuesta + (r?.fueraDeLugar ? ' · fuera de lugar' : '') + ' — listo, escanea el siguiente';
        this.limpiarFotoCaptura();
        c.bien = null;
        c.paso = 'escanear';
        window.scrollTo({ top: 0 });
      },
      error: e => { c.guardando = false; c.error = this.mensajeError(e, 'No se pudo guardar. Revisa tu conexión e inténtalo de nuevo.'); }
    });
  }

  // ----- No encontrado / bien nuevo / traer de otra sede -----

  marcarNoEncontradoCap(): void {
    const c = this.cap, b = c?.bien;
    if (!b || c.guardando) return;
    c.guardando = true;
    c.error = '';
    this.http.post<any>(this.api + '/levantamientos/' + c.lev.idLevantamiento + '/no-encontrados', { idBien: b.id_bien }).subscribe({
      next: () => {
        c.guardando = false;
        b.estatus = 'NO_ENCONTRADO';
        b.respuesta = null;
        b.fuera_de_lugar = false;
        b.fecha_registro = new Date().toISOString();
        this.limpiarFotoCaptura();
        c.bien = null;
        c.paso = 'escanear';
        c.ultimo = b.clave + ' · marcado como no encontrado';
        this.abrirAvanceCaptura('pendientes');
      },
      error: e => { c.guardando = false; c.error = this.mensajeError(e, 'No se pudo marcar. Inténtalo de nuevo.'); }
    });
  }

  abrirNuevoBienCap(): void {
    const c = this.cap;
    if (!c) return;
    this.detenerEscaner();
    const g = c.grupo;
    this.capNuevo = {
      articulo: '', idUbicacion: g?.modo === 'area' ? g.id : null, idPersona: g?.modo === 'persona' ? g.id : null,
      idClasificacion: null, idMarca: null, modelo: '', numeroSerie: '', caracteristicas: '',
      clave: c.claveSugerida || '', mas: !!c.claveSugerida, guardando: false, error: ''
    };
  }

  guardarNuevoBienCap(): void {
    const n = this.capNuevo, c = this.cap;
    if (!n || !c || n.guardando) return;
    if (!String(n.articulo || '').trim()) { n.error = 'Escribe qué es el bien.'; return; }
    if (!n.idUbicacion) { n.error = 'Elige el área donde queda el bien.'; return; }
    n.guardando = true;
    n.error = '';
    const datos = {
      articulo: n.articulo, idUbicacion: n.idUbicacion, idPersona: n.idPersona, idClasificacion: n.idClasificacion,
      idMarca: n.idMarca, modelo: n.modelo, numeroSerie: n.numeroSerie, caracteristicas: n.caracteristicas, clave: n.clave
    };
    this.http.post<any>(this.api + '/levantamientos/' + c.lev.idLevantamiento + '/bienes-nuevos', datos).subscribe({
      next: r => {
        this.capNuevo = null;
        c.bienes.unshift(r);
        this.despuesDeAgregar(r, 'nuevo');
      },
      error: e => { n.guardando = false; n.error = this.mensajeError(e, 'No se pudo registrar el bien.'); }
    });
  }

  private despuesDeAgregar(b: any, tipo: 'nuevo' | 'traido'): void {
    const c = this.cap;
    c.agregado = { bien: b, tipo };
    c.cambios = true;
    c.error = '';
    c.ultimo = '';
    c.ofrecerNuevo = false;
    c.claveSugerida = '';
    c.paso = 'escanear';
    window.scrollTo({ top: 0 });
  }

  abrirBuscarGeneral(q = '', resultados?: any[]): void {
    this.detenerEscaner();
    this.capBuscar = { q, resultados: resultados || [], buscado: !!resultados, cargando: false, elegido: null,
                       idUbicacion: null, idPersona: null, guardando: false, error: '' };
    if (q && !resultados) this.buscarGeneralCap();
  }

  buscarGeneralCap(): void {
    const s = this.capBuscar;
    const q = String(s?.q || '').trim();
    if (!s || s.cargando) return;
    if (q.length < 2) { s.error = 'Escribe al menos 2 letras o números.'; return; }
    s.cargando = true;
    s.error = '';
    s.elegido = null;
    this.http.get<any[]>(this.api + '/levantamientos/' + this.cap.lev.idLevantamiento + '/buscar', { params: { q } }).subscribe({
      next: rows => { s.cargando = false; s.buscado = true; s.resultados = rows || []; },
      error: e => { s.cargando = false; s.error = this.mensajeError(e, 'No se pudo buscar.'); }
    });
  }

  abrirDesdeBusqueda(r: any): void {
    const b = this.cap.bienes.find((x: any) => x.id_bien === r.id_bien) || r;
    this.capBuscar = null;
    this.abrirBienCaptura(b);
  }

  elegirTraer(r: any): void {
    const g = this.cap.grupo;
    this.capBuscar.elegido = r;
    this.capBuscar.idUbicacion = g?.modo === 'area' ? g.id : null;
    this.capBuscar.idPersona = g?.modo === 'persona' ? g.id : null;
    this.capBuscar.error = '';
  }

  confirmarTraer(r: any): void {
    const s = this.capBuscar, c = this.cap;
    if (!s || s.guardando) return;
    if (!s.idUbicacion) { s.error = 'Elige el área donde queda el bien.'; return; }
    s.guardando = true;
    s.error = '';
    this.http.post<any>(this.api + '/levantamientos/' + c.lev.idLevantamiento + '/traer',
      { idBien: r.id_bien, idUbicacion: s.idUbicacion, idPersona: s.idPersona }).subscribe({
      next: b => {
        this.capBuscar = null;
        c.bienes = [b, ...c.bienes.filter((x: any) => x.id_bien !== b.id_bien)];
        this.despuesDeAgregar(b, 'traido');
      },
      error: e => { s.guardando = false; s.error = this.mensajeError(e, 'No se pudo mover el bien.'); }
    });
  }

  nuevoDesdeBusqueda(): void {
    const q = String(this.capBuscar?.q || '').trim();
    this.capBuscar = null;
    this.abrirNuevoBienCap();
    if (/^[A-Za-z]{1,4}[0-9]{2,}$/.test(q)) { this.capNuevo.clave = q.toUpperCase(); this.capNuevo.mas = true; }
    else this.capNuevo.articulo = q;
  }

  // ----- Actas (PDF) -----

  esActa(d: any): boolean {
    return String(d?.tipo || '').startsWith('ACTA');
  }

  tituloDoc(d: any): string {
    switch (d?.tipo) {
      case 'ACTA_INICIO': return 'Acta de inicio';
      case 'ACTA_CIERRE': return 'Acta de cierre';
      case 'BAJA': return 'Formato de Baja';
      case 'RESGUARDO': return 'Resguardo individual';
      default: return 'Formato de Alta';
    }
  }

  etiquetaTipoDoc(d: any): string {
    return d?.tipo === 'ACTA_INICIO' ? 'INICIO' : d?.tipo === 'ACTA_CIERRE' ? 'CIERRE' : d?.tipo;
  }

  get conteoActas(): number {
    return this.documentosLista.filter(d => String(d.tipo).startsWith('ACTA')).length;
  }

  private async crearPdfActa(d: any): Promise<Blob> {
    const doc = new jsPDF({ unit: 'mm', format: 'letter', compress: true });
    const X = 22, W = 172, R = X + W, ALTO = 279.4;
    const logo = await this.cargarLogo();
    const inicio = d.tipo === 'ACTA_INICIO';
    let y = 0;

    const cabecera = () => {
      if (logo) { try { doc.addImage(logo, 'PNG', R - 46, 12, 46, 46 * 227 / 879, 'logo', 'FAST'); } catch { /* sin logo */ } }
      // Pie de página como en el formato oficial
      doc.setFillColor(242, 140, 30); doc.rect(X - 6, 258, 90, 1.6, 'F');
      doc.setFillColor(20, 20, 20); doc.rect(X + 84, 258, R + 6 - (X + 84), 1.6, 'F');
      doc.setFont('helvetica', 'normal'); doc.setFontSize(7.5); doc.setTextColor(40, 40, 40);
      doc.text('COMPROMETIDO CON', X - 4, 264);
      doc.setFont('helvetica', 'bold'); doc.text('LA INFRAESTRUCTURA', X - 4, 267.5);
      doc.setFont('helvetica', 'normal'); doc.text('DE NUESTRO PAÍS', X - 4, 271);
      doc.setFontSize(7); doc.setTextColor(130, 130, 130);
      doc.text('Folio ' + (d.folio || '') + ' · DOZZER', R + 4, 271, { align: 'right' });
      if (d.anulado) this.marcaAgua(doc, 'ANULADO', [194, 65, 75], 150);
      if (d.borrador) this.marcaAgua(doc, 'BORRADOR', [110, 125, 145], 150);
      doc.setTextColor(0, 0, 0);
      y = 42;
    };
    const nuevaPagina = () => { doc.addPage(); cabecera(); };
    const asegurar = (alto: number) => { if (y + alto > 250) nuevaPagina(); };
    const parrafo = (texto: string, opts: { negrita?: boolean; tam?: number; espacio?: number } = {}) => {
      doc.setFont('helvetica', opts.negrita ? 'bold' : 'normal');
      doc.setFontSize(opts.tam || 11);
      const lineas = doc.splitTextToSize(texto, W);
      const alto = lineas.length * (opts.tam || 11) * 0.42;
      asegurar(alto);
      doc.text(lineas, X, y, { maxWidth: W, align: 'justify', lineHeightFactor: 1.2 });
      y += alto + (opts.espacio ?? 5);
    };
    const titulo = (lineas: string[]) => {
      doc.setFont('helvetica', 'bold'); doc.setFontSize(12.5);
      lineas.forEach(l => { doc.text(l, 108, y, { align: 'center' }); y += 6; });
      y += 8;
    };
    const participantes = () => {
      asegurar(30);
      y += 10;
      doc.setFont('helvetica', 'bold'); doc.setFontSize(11);
      doc.text('P a r t i c i p a n t e s', 108, y, { align: 'center' });
      y += 30;
      const ps: any[] = d.participantes || [];
      for (let i = 0; i < ps.length; i += 2) {
        asegurar(34);
        [ps[i], ps[i + 1]].forEach((p, j) => {
          if (!p) return;
          const cx = j === 0 ? X + W * 0.25 : X + W * 0.75;
          doc.setDrawColor(0); doc.setLineWidth(0.3);
          doc.line(cx - 38, y, cx + 38, y);
          doc.setFont('helvetica', 'bold'); doc.setFontSize(10.5);
          doc.text(this.ajustarTexto(doc, p.cargo || '', 80), cx, y + 5, { align: 'center' });
          doc.setFont('helvetica', 'normal');
          doc.text(this.ajustarTexto(doc, p.nombre || '', 80), cx, y + 10, { align: 'center' });
        });
        y += 38;
      }
    };

    cabecera();
    if (inicio) {
      titulo(['ACTA DE INICIO DE LEVANTAMIENTO', 'DE INVENTARIO DE BIENES ' + d.anio]);
      parrafo('En la ciudad de ' + (d.ciudad || '________') + ', al día ' + d.dia + ' del mes de ' + d.mes + ' del año ' + d.anioFecha +
        ' siendo las ' + d.hora + ' hrs, en las instalaciones de la empresa DOZZER, sede ' + d.sede +
        (d.direccion ? ', ubicada en dirección ' + d.direccion : '') +
        ', se da inicio al proceso de levantamiento de inventario de bienes ' + d.anio + ', muebles propiedad de la empresa.');
      parrafo('Objetivo', { negrita: true, espacio: 2 });
      parrafo('Establecer de manera formal el inicio del levantamiento físico de los bienes muebles institucionales, con el propósito de verificar y documentar su existencia, estado de conservación, ubicación real y condiciones de resguardo. Este proceso permitirá asegurar la integridad del patrimonio, actualizar los registros oficiales y generar información confiable para la toma de decisiones administrativas, operativas y de control interno.');
      parrafo('Alcance', { negrita: true, espacio: 2 });
      parrafo('El levantamiento físico de inventario comprende la verificación de todos los bienes muebles que se encuentren en uso, resguardo o almacenamiento dentro de las áreas operativas, administrativas y técnicas.');
      parrafo('Localización y revisión de cada bien mueble en su área correspondiente.', { espacio: 1.5 });
      parrafo('Confirmación de datos clave:', { espacio: 1.5 });
      ['Identificación del bien', 'Número de inventario', 'Estado de conservación', 'Ubicación física', 'Responsable o área de resguardo']
        .forEach(t => parrafo('•  ' + t, { espacio: 1 }));
      y += 4;
      parrafo('Con lo anterior, y no habiendo más asuntos que tratar, se da por formalmente iniciado el proceso de levantamiento físico de inventario de bienes muebles, quedando asentado que las actividades se realizarán conforme a la metodología, alcances y responsabilidades previamente establecidos.');
      nuevaPagina();
      parrafo('Las partes involucradas manifiestan su conformidad con el contenido de la presente acta y se comprometen a colaborar en el desarrollo del proceso, proporcionando la información, facilidades y acceso necesarios para su adecuada ejecución.');
      parrafo('Para constancia y a fin de que surta los efectos administrativos correspondientes, se firma la presente Acta de Inicio en la fecha y lugar señalados, levantándose el documento en 2 tantos, quedando un ejemplar bajo resguardo del área responsable del control patrimonial.');
      participantes();
    } else {
      const t = d.totales || {};
      titulo(['ACTA DE CIERRE DE INVENTARIO ' + d.anio]);
      parrafo('Siendo las ' + d.hora + ' hrs. del día ' + d.dia + ' de ' + d.mes + ' del ' + d.anioFecha +
        ', se declara formalmente concluido el levantamiento físico de inventario de bienes muebles correspondiente a DOZZER, sede ' + d.sede +
        (d.direccion ? ', ubicada en dirección ' + d.direccion : '') +
        '. Durante este proceso se verificó la existencia, estado de conservación y ubicación de los bienes revisados, realizándose las actividades conforme a la metodología, alcances y lineamientos establecidos en el Acta de Inicio' +
        (d.fechaInicioTexto ? ' del ' + d.fechaInicioTexto : '') + '.');
      parrafo('HECHOS', { negrita: true, espacio: 2 });
      parrafo('Las áreas participantes manifiestan su conformidad con los resultados obtenidos y con la información registrada en los sistemas institucionales, reconociendo que los datos generados servirán como base para la actualización del inventario oficial y para la toma de decisiones administrativas, operativas y de control interno.');
      parrafo('Resumen de Inventario', { negrita: true, espacio: 4 });
      const resumen: any[] = d.resumen || [];
      doc.setFont('helvetica', 'bold'); doc.setFontSize(9.5);
      const colW = W / 3;
      for (let i = 0; i < resumen.length; i += 3) {
        asegurar(7);
        resumen.slice(i, i + 3).forEach((c, j) => {
          // El número siempre visible; el nombre se acorta si no cabe
          const num = String(c.total);
          const anchoNum = doc.getTextWidth(num) + 2;
          doc.text(this.ajustarTexto(doc, this.capitalizar(c.clasificacion), colW - 4 - anchoNum), X + 2 + j * colW, y);
          doc.text(num, X + j * colW + colW - 3, y, { align: 'right' });
        });
        y += 6.5;
      }
      y += 4;
      const respuestas = Object.entries(t.respuestas || {}).map(([k, v]) => k + ': ' + v).join(' · ');
      parrafo('Bienes registrados en la sede: ' + (t.bienes ?? 0) + '. Localizados: ' + (t.encontrados ?? 0) + '. No localizados: ' + (t.noLocalizados ?? 0) +
        '. Fuera de su ubicación registrada: ' + (t.fueraDeLugar ?? 0) + '.', { tam: 10, espacio: 2 });
      if (respuestas) parrafo((d.pregunta || 'Estado') + ' — ' + respuestas + '.', { tam: 10, espacio: 2 });
      const extras: string[] = [];
      if (t.nuevos) extras.push('bienes nuevos registrados durante el levantamiento: ' + t.nuevos);
      if (t.transferidos) extras.push('bienes traídos de otra sede: ' + t.transferidos);
      if (t.marcadosNoEncontrados) extras.push('marcados como no encontrados: ' + t.marcadosNoEncontrados);
      if (extras.length) parrafo(this.capitalizarPrimera(extras.join('; ')) + '.', { tam: 10 });
      if (d.borrador) parrafo('Borrador generado con la información capturada hasta este momento.', { tam: 9 });
      parrafo('Se adjunta listado de bienes.', { negrita: true });
      nuevaPagina();
      parrafo('No habiendo más asuntos que tratar, se procede a la formalización de la presente Acta de Cierre, la cual se firma en 2 tantos, destinándose un ejemplar al área responsable del control patrimonial para los efectos administrativos y normativos correspondientes.');
      participantes();

      // Anexo: listado de bienes
      const listado: any[] = d.listado || [];
      if (listado.length) {
        nuevaPagina();
        doc.setFont('helvetica', 'bold'); doc.setFontSize(11.5);
        doc.text('ANEXO · LISTADO DE BIENES', X, y); y += 7;
        const cols = [10, 26, 70, 38, 28];
        const encabezado = () => {
          doc.setFillColor(23, 38, 61); doc.rect(X, y - 4.5, W, 7, 'F');
          doc.setFont('helvetica', 'bold'); doc.setFontSize(8); doc.setTextColor(255, 255, 255);
          let cx = X;
          ['#', 'Clave', 'Descripción', 'Ubicación', 'Resultado'].forEach((h, i) => { doc.text(h, cx + 1.5, y); cx += cols[i]; });
          doc.setTextColor(0, 0, 0); y += 6;
        };
        encabezado();
        listado.forEach((f, n) => {
          if (y > 248) { nuevaPagina(); encabezado(); }
          if (n % 2 === 0) { doc.setFillColor(244, 246, 250); doc.rect(X, y - 4, W, 6, 'F'); }
          doc.setFont('helvetica', 'normal'); doc.setFontSize(7.8);
          const resultado = f.encontrado ? (f.respuesta || 'Localizado') + (f.fueraDeLugar ? ' *' : '') : 'No localizado';
          const celdas = [String(n + 1), f.clave, f.descripcion, f.ubicacion, resultado];
          let cx = X;
          celdas.forEach((c, i) => {
            if (i === 4 && !f.encontrado) doc.setTextColor(194, 65, 75);
            doc.text(this.ajustarTexto(doc, String(c || ''), cols[i] - 3), cx + 1.5, y);
            doc.setTextColor(0, 0, 0);
            cx += cols[i];
          });
          y += 6;
        });
        y += 3;
        doc.setFont('helvetica', 'italic'); doc.setFontSize(7.5); doc.setTextColor(110, 110, 110);
        doc.text('* Localizado fuera de su ubicación registrada.', X, y);
      }
    }
    return doc.output('blob');
  }

  private marcaAgua(doc: jsPDF, texto: string, color: [number, number, number], y: number): void {
    try { (doc as any).setGState(new (doc as any).GState({ opacity: 0.16 })); } catch { /* sin transparencia */ }
    doc.setFont('helvetica', 'bold'); doc.setFontSize(texto.length > 8 ? 62 : 72); doc.setTextColor(...color);
    doc.text(texto, 108, y, { align: 'center', angle: 32 });
    try { (doc as any).setGState(new (doc as any).GState({ opacity: 1 })); } catch { /* sin transparencia */ }
    doc.setTextColor(0, 0, 0);
  }

  private capitalizarPrimera(t: string): string {
    return t ? t.charAt(0).toUpperCase() + t.slice(1) : t;
  }

  private ajustarTexto(doc: jsPDF, t: string, ancho: number): string {
    t = String(t || '');
    if (doc.getTextWidth(t) <= ancho) return t;
    while (t.length && doc.getTextWidth(t + '...') > ancho) t = t.slice(0, -1);
    return t + '...';
  }

  private capitalizar(t: string): string {
    return String(t || '').toLowerCase().replace(/(^|\s)\S/g, x => x.toUpperCase());
  }

  // ===== Documentos guardados =====

  loadDocumentos(): void {
    if (!this.logged) return;
    this.docsCargando = true;
    this.docsError = '';
    this.http.get<any[]>(this.api + '/documentos').subscribe({
      next: rows => { this.documentosLista = rows || []; this.docsCargando = false; },
      error: e => { this.docsCargando = false; this.docsError = this.mensajeError(e, 'No fue posible consultar los documentos.'); }
    });
  }

  conteoDocs(tipo: string): number {
    return this.documentosLista.filter(d => d.tipo === tipo && d.estado !== 'ANULADO').length;
  }

  get documentosFiltrados(): any[] {
    const palabras = this.docsBusqueda.toLowerCase().trim().split(/\s+/).filter(Boolean);
    return this.documentosLista
      .filter(d => this.docsVista === 'ACTAS' ? String(d.tipo).startsWith('ACTA') : d.tipo === this.docsVista)
      .filter(d => {
        if (!palabras.length) return true;
        const t = [d.folio, d.clave, d.descripcion, d.sede, d.solicito, d.autorizo, d.persona, d.levantamiento].join(' ').toLowerCase();
        return palabras.every(p => t.includes(p));
      });
  }

  abrirDocumentoGuardado(id: any): void {
    if (!id || this.abriendoDocumento) return;
    this.abriendoDocumento = id;
    this.docGenerando = true;
    this.http.get<any>(this.api + '/documentos/' + id).subscribe({
      next: r => {
        this.abriendoDocumento = null;
        this.docGenerando = false;
        this.abrirFormato({ ...(r.datos || {}), anulado: r.estado === 'ANULADO', borrador: r.estado === 'BORRADOR',
                            historico: r.estado === 'HISTORICO' } as FormatoDoc);
      },
      error: e => {
        this.abriendoDocumento = null;
        this.docGenerando = false;
        this.notificar('error', 'No se abrió el documento', this.mensajeError(e, 'Inténtalo de nuevo.'), 0);
      }
    });
  }

  // Abre el documento guardado; si no existe (bienes anteriores), lo arma con los datos actuales
  verOFormato(idDocumento: any, generar: () => FormatoDoc): void {
    if (idDocumento) this.abrirDocumentoGuardado(idDocumento);
    else this.abrirFormato(generar());
  }

  private buscarDocumento(params: string, generar: () => FormatoDoc): void {
    this.http.get<any[]>(this.api + '/documentos?' + params).subscribe({
      next: rows => {
        const doc = (rows || []).find(d => d.estado !== 'ANULADO') || (rows || [])[0];
        this.verOFormato(doc?.idDocumento, generar);
      },
      error: () => this.abrirFormato(generar())
    });
  }

  formatoAltaDeBien(b: any): void {
    this.buscarDocumento('tipo=ALTA&idBien=' + b.id_bien, () =>
      this.formatoDesdeBien('ALTA', b, { folio: this.folio('ALTA', null, b.id_bien), solicito: this.miNombre, autorizo: this.miNombre }));
  }

  formatoDeSolicitud(s: any): void {
    this.buscarDocumento('idSolicitud=' + s.idSolicitud, () =>
      this.formatoDesdeSolicitud(s, s.idBien ? this.bienPorId(s.idBien) : null));
  }

  async abrirFormato(datos: FormatoDoc): Promise<void> {
    this.cerrarFormato();
    this.docGenerando = true;
    this.docError = '';
    this.docMensaje = 'Se adjunta: ' + this.tituloDoc(datos) + ' con folio ' + datos.folio + '.';
    this.waAviso = '';
    this.vistaDoc = 'inicio';
    try {
      const blob = await this.crearPdfFormato(datos);
      const url = URL.createObjectURL(blob);
      this.documento = {
        datos, url, blob,
        vista: this.sanitizer.bypassSecurityTrustResourceUrl(url),
        nombre: this.tituloDoc(datos) + ' ' + datos.folio + '.pdf'
      };
    } catch {
      this.notificar('error', 'No se generó el formato', 'Ocurrió un problema al crear el PDF. Inténtalo de nuevo.', 0);
    } finally {
      this.docGenerando = false;
    }
  }

  cerrarFormato(): void {
    if (this.documento) URL.revokeObjectURL(this.documento.url);
    this.documento = null;
  }

  descargarFormato(): void {
    if (!this.documento) return;
    const a = document.createElement('a');
    a.href = this.documento.url;
    a.download = this.documento.nombre;
    document.body.appendChild(a);
    a.click();
    a.remove();
  }

  imprimirFormato(): void {
    if (!this.documento) return;
    const w = window.open(this.documento.url, '_blank');
    if (!w) this.descargarFormato();
  }

  abrirVistaDoc(vista: 'whatsapp'): void {
    this.vistaDoc = vista;
    this.docError = '';
    this.waAviso = '';
    this.waUrl = '';
  }

  // Abre WhatsApp (app en celular, WhatsApp Web/escritorio en PC) con el mensaje listo y descarga el PDF para adjuntarlo.
  // WhatsApp no acepta archivos por enlace, por eso el PDF se adjunta con el clip.
  abrirWhatsApp(): void {
    const d = this.documento;
    if (!d) return;
    let numero = (this.waNumero || '').replace(/\D/g, '');
    if (numero && numero.length !== 10 && numero.length !== 12 && numero.length !== 13) {
      this.waAviso = 'Escribe el número a 10 dígitos (ej. 4611234567) o déjalo vacío para elegir el contacto en WhatsApp.';
      return;
    }
    if (numero.length === 10) numero = '52' + numero; // México
    const resg = d.datos.tipo === 'RESGUARDO';
    const bien: any = resg ? { inventario: (d.datos as any).persona?.nombre, descripcion: (d.datos as any).totalBienes + ' bienes' } : (d.datos.filas || [])[0];
    const texto = this.tituloDoc(d.datos) + ' · Folio ' + d.datos.folio +
      (bien ? '\n' + [bien.inventario, bien.descripcion].filter(x => x).join(' · ') : '') +
      (d.datos.sede ? '\nSede: ' + d.datos.sede : '') +
      '\n\nTe comparto el PDF del formato (va adjunto).';
    // Celular: abre la app. Computadora: WhatsApp Web directo (sin la página intermedia de wa.me)
    const movil = /Android|iPhone|iPad|iPod|Mobile/i.test(navigator.userAgent);
    const params = (numero ? 'phone=' + numero + '&' : '') + 'text=' + encodeURIComponent(texto);
    const url = (movil ? 'https://api.whatsapp.com/send?' : 'https://web.whatsapp.com/send?') + params;
    this.waUrl = url;

    // Se abre con un enlace real en pestaña nueva; nunca se sale de DOZZER
    const a = document.createElement('a');
    a.href = url;
    a.target = '_blank';
    a.rel = 'noopener';
    document.body.appendChild(a);
    a.click();
    a.remove();

    // La descarga va un instante después para que el navegador no cancele la apertura
    setTimeout(() => this.descargarFormato(), 700);
    this.waAviso = 'Se descargó "' + d.nombre + '". En el chat de WhatsApp toca 📎 → Documento y adjúntalo.';
  }

  async compartirFormato(): Promise<void> {
    if (!this.documento) return;
    try {
      const archivo = new File([this.documento.blob], this.documento.nombre, { type: 'application/pdf' });
      await (navigator as any).share({ files: [archivo], title: this.documento.nombre, text: this.docMensaje });
    } catch { /* el usuario canceló */ }
  }

  private async cargarLogo(): Promise<string | null> {
    if (this.logoDataUrl) return this.logoDataUrl;
    try {
      const resp = await fetch('logo-dozzer.png');
      if (!resp.ok) return null;
      const blob = await resp.blob();
      this.logoDataUrl = await new Promise<string>((ok, mal) => {
        const fr = new FileReader();
        fr.onload = () => ok(String(fr.result));
        fr.onerror = mal;
        fr.readAsDataURL(blob);
      });
      return this.logoDataUrl;
    } catch { return null; }
  }

  // Formato de Alta/Baja — diseño "Institucional moderno" (propuesta A elegida).
  // Contenido guiado por los formatos oficiales FA-04-02-03 / FA-04-02-04.
  // Las firmas quedan como espacio vacío: ahí irá la imagen de la firma de cada persona.
  private async crearPdfFormato(d: FormatoDoc): Promise<Blob> {
    if (this.esActa(d)) return this.crearPdfActa(d as any);
    if (d.tipo === 'RESGUARDO') return this.crearPdfResguardo(d as any);
    type RGB = [number, number, number];
    const NAVY: RGB = [23, 38, 61], AZUL: RGB = [36, 76, 120], GRIS: RGB = [110, 120, 135];
    const CLARO: RGB = [244, 246, 250], LINEA: RGB = [214, 220, 229];
    const ACENTO: RGB = d.tipo === 'ALTA' ? [33, 131, 90] : [194, 65, 75];
    const doc = new jsPDF({ unit: 'mm', format: 'letter', compress: true });
    const X = 20, W = 176, R = X + W;
    const ajustar = (t: string, ancho: number) => {
      t = String(t || '');
      if (doc.getTextWidth(t) <= ancho) return t;
      while (t.length && doc.getTextWidth(t + '...') > ancho) t = t.slice(0, -1);
      return t + '...';
    };
    const seccion = (titulo: string, y: number, color: RGB = AZUL) => {
      doc.setFont('helvetica', 'bold'); doc.setFontSize(9); doc.setTextColor(...color);
      doc.text(titulo, X, y);
    };

    // Encabezado
    const logo = await this.cargarLogo();
    if (logo) {
      try { doc.addImage(logo, 'PNG', X, 15, 52, 52 * 227 / 879, 'logo', 'FAST'); } catch { /* sin logo */ }
    }
    doc.setFont('helvetica', 'bold'); doc.setFontSize(17); doc.setTextColor(...NAVY);
    doc.text('FORMATO DE ' + d.tipo, R, 21, { align: 'right' });
    doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor(...GRIS);
    doc.text('Clave ' + (d.tipo === 'ALTA' ? 'FA-04-02-03' : 'FA-04-02-04') + '   ·   Revisión 0   ·   Emisión 01/06/2026', R, 27, { align: 'right' });
    doc.setDrawColor(...NAVY); doc.setLineWidth(0.8); doc.line(X, 34, R, 34);
    doc.setDrawColor(...ACENTO); doc.setLineWidth(0.8); doc.line(R - 30, 34, R, 34);

    // Datos generales
    const cajas: [string, string][] = [['SEDE', d.sede || '—'], ['FOLIO', d.folio], ['FECHA DE SOLICITUD', d.fecha]];
    const anchosCaja = [50, 46, 76];
    let x = X;
    cajas.forEach(([l, v], i) => {
      doc.setFillColor(...CLARO); doc.roundedRect(x, 40, anchosCaja[i], 15, 2, 2, 'F');
      doc.setFont('helvetica', 'bold'); doc.setFontSize(7); doc.setTextColor(...GRIS); doc.text(l, x + 4, 45.5);
      doc.setFontSize(11); doc.setTextColor(...NAVY); doc.text(ajustar(v, anchosCaja[i] - 8), x + 4, 51.5);
      x += anchosCaja[i] + 2;
    });

    // Tabla de bienes (solo las filas con datos)
    seccion(d.tipo === 'ALTA' ? 'BIENES QUE SE DAN DE ALTA' : 'BIENES QUE SE DAN DE BAJA', 66);
    const cols = [14, 30, 82, 50];
    const titulos = ['Núm.', 'Inventario', 'Descripción del bien', 'Modelo / Serie'];
    doc.setFillColor(...NAVY); doc.rect(X, 69, W, 8, 'F');
    doc.setFont('helvetica', 'bold'); doc.setFontSize(8); doc.setTextColor(255, 255, 255);
    x = X; titulos.forEach((t, i) => { doc.text(t, x + 3, 74.2); x += cols[i]; });
    let y = 77;
    doc.setFont('helvetica', 'normal');
    d.filas.forEach((f, n) => {
      if (n % 2 === 0) { doc.setFillColor(...CLARO); doc.rect(X, y, W, 10, 'F'); }
      doc.setFontSize(8.5); doc.setTextColor(...NAVY);
      const celdas = [String(n + 1), f.inventario, f.descripcion, f.modeloSerie];
      x = X;
      celdas.forEach((t, i) => {
        const ancho = cols[i] - 5;
        if (doc.getTextWidth(String(t || '')) <= ancho || i < 2) {
          doc.text(ajustar(String(t || ''), ancho), x + 3, y + 6);
        } else {
          doc.setFontSize(7);
          const renglones = doc.splitTextToSize(String(t), ancho);
          if (renglones.length === 1) doc.text(renglones[0], x + 3, y + 6);
          else doc.text([renglones[0], ajustar(renglones.slice(1).join(' '), ancho)], x + 3, y + 4.2, { lineHeightFactor: 1.15 });
          doc.setFontSize(8.5);
        }
        x += cols[i];
      });
      y += 10;
    });
    doc.setDrawColor(...LINEA); doc.setLineWidth(0.3); doc.line(X, y, R, y);

    // Detalle del bien (cuando es uno solo)
    const detalle = (d.detalle || []).filter(c => c.valor && c.valor.trim());
    y += 8;
    if (d.filas.length === 1 && detalle.length) {
      detalle.forEach((c, i) => {
        const cx = X + (i % 3) * 59, cy = y + Math.floor(i / 3) * 12;
        doc.setFont('helvetica', 'bold'); doc.setFontSize(6.8); doc.setTextColor(...GRIS); doc.text(c.etiqueta.toUpperCase(), cx, cy);
        doc.setFont('helvetica', 'normal'); doc.setFontSize(8.5); doc.setTextColor(...NAVY); doc.text(ajustar(c.valor, 56), cx, cy + 4.6);
      });
      y += Math.ceil(detalle.length / 3) * 12 + 6;
    }

    // Observaciones
    seccion(d.tipo === 'ALTA' ? 'OBSERVACIONES' : 'MOTIVO Y OBSERVACIONES', y);
    doc.setFillColor(...CLARO); doc.roundedRect(X, y + 3, W, 26, 2, 2, 'F');
    doc.setFont('helvetica', 'normal'); doc.setFontSize(8.5); doc.setTextColor(...NAVY);
    const lineas = doc.splitTextToSize(d.observaciones || 'Sin observaciones.', W - 8).slice(0, 5);
    doc.text(lineas, X + 4, y + 9.5);
    y += 3 + 26 + 13;

    // Firmas: espacio vacío para la imagen de la firma, línea, nombre y rol
    seccion('FIRMAS', y);
    const espacioFirma = (fx: number, fy: number, w: number, nombre: string, rol: string) => {
      doc.setDrawColor(...LINEA); doc.setLineWidth(0.3);
      (doc as any).setLineDashPattern([1.2, 1.2], 0);
      doc.roundedRect(fx + 4, fy, w - 8, 22, 2, 2);
      (doc as any).setLineDashPattern([], 0);
      doc.setFont('helvetica', 'italic'); doc.setFontSize(7); doc.setTextColor(170, 178, 190);
      doc.text('Espacio para firma', fx + w / 2, fy + 12, { align: 'center' });
      doc.setDrawColor(...NAVY); doc.setLineWidth(0.35);
      doc.line(fx + 4, fy + 26, fx + w - 4, fy + 26);
      if (nombre) {
        doc.setFont('helvetica', 'bold'); doc.setFontSize(8); doc.setTextColor(...NAVY);
        doc.text(ajustar(nombre.toUpperCase(), w - 6), fx + w / 2, fy + 30.5, { align: 'center' });
      } else {
        doc.setFont('helvetica', 'italic'); doc.setFontSize(7.5); doc.setTextColor(170, 178, 190);
        doc.text('Nombre y cargo', fx + w / 2, fy + 30.5, { align: 'center' });
      }
      doc.setFont('helvetica', 'normal'); doc.setFontSize(7.5); doc.setTextColor(...GRIS);
      doc.text(rol, fx + w / 2, fy + 34.5, { align: 'center' });
    };
    const fy = Math.min(y + 5, 222);
    espacioFirma(X, fy, 58, d.solicito, 'Solicitó');
    espacioFirma(X + 59, fy, 58, d.autorizo, 'Autorizó');
    // Valida siempre el Jefe de recursos materiales (alta y baja)
    espacioFirma(X + 118, fy, 58, 'Ing. Juan Carlos Zuno Muñiz', 'Validó · Jefe de recursos materiales');

    if (d.anulado) {
      try { (doc as any).setGState(new (doc as any).GState({ opacity: 0.18 })); } catch { /* sin transparencia */ }
      doc.setFont('helvetica', 'bold'); doc.setFontSize(72); doc.setTextColor(194, 65, 75);
      doc.text('ANULADO', 108, 165, { align: 'center', angle: 32 });
      try { (doc as any).setGState(new (doc as any).GState({ opacity: 1 })); } catch { /* sin transparencia */ }
    }

    // Pie de página
    doc.setDrawColor(...LINEA); doc.setLineWidth(0.3); doc.line(X, 264, R, 264);
    doc.setFont('helvetica', 'normal'); doc.setFontSize(7); doc.setTextColor(...GRIS);
    doc.text('DOZZER Construcción y Asesoría S.A. de C.V.', X, 269);
    const ahora = new Date();
    const sello = ahora.toLocaleDateString('es-MX', { day: '2-digit', month: '2-digit', year: 'numeric' }) + ' ' +
      ahora.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit', hour12: false });
    doc.text('Folio ' + d.folio + ' · Generado en DOZZER el ' + sello, R, 269, { align: 'right' });

    return doc.output('blob');
  }

  // Resguardo individual FA-04-02-02 (mismo estilo que el Alta). Lista los bienes REVISADOS de la persona.
  private async crearPdfResguardo(d: any): Promise<Blob> {
    type RGB = [number, number, number];
    const NAVY: RGB = [23, 38, 61], AZUL: RGB = [36, 76, 120], GRIS: RGB = [110, 120, 135];
    const CLARO: RGB = [244, 246, 250], LINEA: RGB = [214, 220, 229], ACENTO: RGB = [106, 79, 179];
    const doc = new jsPDF({ unit: 'mm', format: 'letter', compress: true });
    const X = 20, W = 176, R = X + W, LIMITE = 252;
    const per = d.persona || {};
    const filas: any[] = d.filas || [];
    const ajustar = (t: any, ancho: number) => this.ajustarTexto(doc, String(t ?? ''), ancho);
    const seccion = (titulo: string, y: number) => {
      doc.setFont('helvetica', 'bold'); doc.setFontSize(9); doc.setTextColor(...AZUL);
      doc.text(titulo, X, y);
    };
    const logo = await this.cargarLogo();

    const encabezado = (completo: boolean) => {
      if (logo) { try { doc.addImage(logo, 'PNG', X, 15, completo ? 52 : 40, (completo ? 52 : 40) * 227 / 879, 'logo', 'FAST'); } catch { /* sin logo */ } }
      doc.setFont('helvetica', 'bold'); doc.setFontSize(completo ? 16 : 12.5); doc.setTextColor(...NAVY);
      doc.text('RESGUARDO DE INVENTARIOS', R, 21, { align: 'right' });
      doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor(...GRIS);
      doc.text('Clave ' + (d.formato || 'FA-04-02-02') + '   ·   Revisión ' + (d.revision || '2') + (completo ? '' : '   ·   ' + (per.nombre || '')), R, 27, { align: 'right' });
      doc.setDrawColor(...NAVY); doc.setLineWidth(0.8); doc.line(X, 34, R, 34);
      doc.setDrawColor(...ACENTO); doc.setLineWidth(0.8); doc.line(R - 30, 34, R, 34);
    };

    // Columnas de la tabla (suman W)
    const cols = [7, 36, 34, 19, 29, 24, 10, 17];
    const titulos = ['#', 'Nombre', 'Descripción', 'Clave', 'Modelo / Serie', 'Área', 'Pzas', 'Estado'];
    const cabeceraTabla = (y: number) => {
      doc.setFillColor(...NAVY); doc.rect(X, y, W, 7.5, 'F');
      doc.setFont('helvetica', 'bold'); doc.setFontSize(7.5); doc.setTextColor(255, 255, 255);
      let x = X;
      titulos.forEach((t, i) => { doc.text(t, x + 1.8, y + 5); x += cols[i]; });
      return y + 7.5;
    };

    // Página 1: datos generales y del usuario
    encabezado(true);
    const cajas: [string, string][] = [['SEDE', d.sede || '—'], ['FOLIO', d.folio || '—'], ['FECHA DEL RESGUARDO', d.fecha || '—']];
    const anchosCaja = [62, 46, 64];
    let x = X;
    cajas.forEach(([l, v], i) => {
      doc.setFillColor(...CLARO); doc.roundedRect(x, 40, anchosCaja[i], 15, 2, 2, 'F');
      doc.setFont('helvetica', 'bold'); doc.setFontSize(7); doc.setTextColor(...GRIS); doc.text(l, x + 4, 45.5);
      doc.setFontSize(11); doc.setTextColor(...NAVY); doc.text(ajustar(v, anchosCaja[i] - 8), x + 4, 51.5);
      x += anchosCaja[i] + 2;
    });

    seccion('DATOS DEL USUARIO', 64);
    const datos: [string, string][] = [
      ['Número del empleado', per.numeroEmpleado || '—'], ['Nombre del empleado', per.nombre || '—'],
      ['Área', per.area || '—'], ['Cargo o puesto', per.puesto || '—']
    ];
    datos.forEach(([l, v], i) => {
      const cx = X + (i % 2) * 89, cy = 71 + Math.floor(i / 2) * 12;
      doc.setFont('helvetica', 'bold'); doc.setFontSize(6.8); doc.setTextColor(...GRIS); doc.text(l.toUpperCase(), cx, cy);
      doc.setFont('helvetica', 'normal'); doc.setFontSize(9.5); doc.setTextColor(...NAVY); doc.text(ajustar(v, 86), cx, cy + 5);
    });

    seccion('BIENES BAJO RESGUARDO (' + filas.length + ')' + (d.levantamiento ? '  ·  ' + d.levantamiento : ''), 100);
    let y = cabeceraTabla(103);
    const marca = (f: any) => f.origen === 'NUEVO' ? ' *' : f.origen === 'TRANSFERIDO' ? ' **' : '';
    filas.forEach((f, n) => {
      if (y + 7 > LIMITE) {
        doc.addPage();
        encabezado(false);
        y = cabeceraTabla(40);
      }
      if (n % 2 === 0) { doc.setFillColor(...CLARO); doc.rect(X, y, W, 7, 'F'); }
      doc.setFont('helvetica', 'normal'); doc.setFontSize(7.4); doc.setTextColor(...NAVY);
      const celdas = [String(n + 1), f.nombre, f.descripcion || '—', String(f.clave || '') + marca(f), f.modeloSerie, f.area, String(f.piezas ?? 1), f.estado || '—'];
      x = X;
      celdas.forEach((t, i) => {
        if (i === 3) doc.setFont('helvetica', 'bold');
        doc.text(ajustar(t, cols[i] - 3), x + 1.8, y + 4.8);
        if (i === 3) doc.setFont('helvetica', 'normal');
        x += cols[i];
      });
      y += 7;
    });
    doc.setDrawColor(...LINEA); doc.setLineWidth(0.3); doc.line(X, y, R, y);
    y += 5;
    if (filas.some(f => f.origen)) {
      doc.setFont('helvetica', 'italic'); doc.setFontSize(7.2); doc.setTextColor(...GRIS);
      doc.text('*  Bien nuevo registrado en el levantamiento.    **  Bien traído de otra sede durante el levantamiento.', X, y);
      y += 5;
    }

    // Firmas (si no caben, van en una hoja nueva)
    if (y + 50 > LIMITE + 8) { doc.addPage(); encabezado(false); y = 44; }
    y += 4;
    seccion('FIRMAS', y);
    const espacioFirma = (fx: number, fy: number, w: number, nombre: string, rol: string) => {
      doc.setDrawColor(...LINEA); doc.setLineWidth(0.3);
      (doc as any).setLineDashPattern([1.2, 1.2], 0);
      doc.roundedRect(fx + 4, fy, w - 8, 22, 2, 2);
      (doc as any).setLineDashPattern([], 0);
      doc.setFont('helvetica', 'italic'); doc.setFontSize(7); doc.setTextColor(170, 178, 190);
      doc.text('Espacio para firma', fx + w / 2, fy + 12, { align: 'center' });
      doc.setDrawColor(...NAVY); doc.setLineWidth(0.35);
      doc.line(fx + 4, fy + 26, fx + w - 4, fy + 26);
      doc.setFont('helvetica', 'bold'); doc.setFontSize(8); doc.setTextColor(...NAVY);
      doc.text(ajustar(String(nombre || '').toUpperCase(), w - 6), fx + w / 2, fy + 30.5, { align: 'center' });
      doc.setFont('helvetica', 'normal'); doc.setFontSize(7.5); doc.setTextColor(...GRIS);
      doc.text(rol, fx + w / 2, fy + 34.5, { align: 'center' });
    };
    espacioFirma(X + 4, y + 5, 80, per.nombre || '', 'Nombre y firma del resguardante');
    espacioFirma(X + 92, y + 5, 80, 'Ing. ' + (d.recursosNombre || 'Juan Carlos Zuno Muñiz'), 'Nombre y firma del ' + String(d.recursosCargo || 'Jefe de recursos materiales').toLowerCase());

    // Marca de agua y pie en todas las hojas
    const paginas = doc.getNumberOfPages();
    const ahora = new Date();
    const sello = ahora.toLocaleDateString('es-MX', { day: '2-digit', month: '2-digit', year: 'numeric' }) + ' ' +
      ahora.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit', hour12: false });
    for (let i = 1; i <= paginas; i++) {
      doc.setPage(i);
      if (d.borrador) this.marcaAgua(doc, 'BORRADOR', [110, 125, 145], 165);
      if (d.historico) this.marcaAgua(doc, 'HISTÓRICO', [110, 125, 145], 165);
      doc.setDrawColor(...LINEA); doc.setLineWidth(0.3); doc.line(X, 264, R, 264);
      doc.setFont('helvetica', 'normal'); doc.setFontSize(7); doc.setTextColor(...GRIS);
      doc.text('DOZZER Construcción y Asesoría S.A. de C.V.', X, 269);
      doc.text('Folio ' + (d.folio || '') + ' · Página ' + i + ' de ' + paginas + ' · Generado el ' + sello, R, 269, { align: 'right' });
    }
    return doc.output('blob');
  }

  // Ventana de aviso: se cierra sola (ms > 0) o con "Entendido"; los errores esperan al usuario (ms = 0)
  notificar(tipo: 'exito' | 'pendiente' | 'baja' | 'rechazo' | 'error' | 'info', titulo: string, mensaje: string, ms = 3500,
            accion?: { texto: string; fn: () => void }): void {
    clearTimeout(this.avisoTimer);
    if (accion) ms = 0; // con acción, espera a que el usuario decida
    this.aviso = { tipo, titulo, mensaje, ms, accion };
    if (ms > 0) this.avisoTimer = setTimeout(() => this.aviso = null, ms);
  }

  cerrarAviso(): void {
    clearTimeout(this.avisoTimer);
    this.aviso = null;
  }

  abrirDialogo(opciones: DialogoConfirmacion): void {
    this.dialogo = { ...opciones, cargando: false, error: '' };
  }

  cerrarDialogo(): void {
    if (this.dialogo?.cargando) return;
    this.dialogo = null;
  }

  // Ejecuta la acción; el diálogo se queda abierto (cargando) hasta que termine, y muestra el error si falla
  confirmarDialogo(): void {
    const d = this.dialogo;
    if (!d || d.cargando) return;
    d.cargando = true;
    d.error = '';
    d.accion(error => {
      if (this.dialogo !== d) return;
      d.cargando = false;
      if (error) d.error = error;
      else this.dialogo = null;
    });
  }

  // Campos del bien que cambian respecto al registro original (para mostrarlos antes de guardar)
  private cambiosBien(): { etiqueta: string; actual: string; propuesto: string }[] {
    const original = this.bienOriginal;
    if (!original) return [];
    return this.camposBien
      .map(c => ({
        etiqueta: c.etiqueta,
        actual: this.valorLegible(c, original[c.claveActual]),
        propuesto: this.valorLegible(c, this.bien[c.clave])
      }))
      .filter(c => c.actual !== c.propuesto);
  }

  // Al editar se pide confirmación mostrando los cambios; un registro nuevo se guarda directo
  saveBien(): void {
    if (this.saving) return;
    if (!this.bien.idBien) {
      this.guardarBien();
      return;
    }

    const cambios = this.cambiosBien();
    if (!cambios.length) {
      this.editing = false;
      this.notificar('info', 'Sin cambios', 'No modificaste ningún dato del bien.');
      return;
    }

    const camposDeEtiqueta = ['Clave', 'Nombre del bien', 'Marca', 'Modelo', 'Número de serie', 'Clasificación', 'Ubicación', 'Responsable de área'];
    this.abrirDialogo({
      tipo: 'confirmar',
      icono: '✎',
      titulo: this.isAdmin ? '¿Guardar los cambios?' : '¿Enviar la solicitud de cambio?',
      mensaje: this.isAdmin
        ? 'Los cambios se aplicarán de inmediato en el inventario.'
        : 'Tus cambios se enviarán al administrador y se aplicarán solo si los autoriza.',
      textoConfirmar: this.isAdmin ? 'Guardar cambios' : 'Enviar solicitud',
      textoCancelar: 'Seguir editando',
      bien: this.bienOriginal,
      cambios,
      nota: cambios.some(c => camposDeEtiqueta.includes(c.etiqueta))
        ? 'Cambiaste datos que aparecen en la etiqueta QR. Reimprime la etiqueta para que muestre la información actualizada.'
        : undefined,
      accion: terminar => this.guardarBien(terminar)
    });
  }

  private guardarBien(terminar?: (error?: string) => void): void {
    this.saving = true;

    const esNuevo = !this.bien.idBien;
    const formulario = { ...this.bien };
    const req = this.bien.idBien
      ? this.http.put(this.api + '/bienes/' + this.bien.idBien, this.bien)
      : this.http.post(this.api + '/bienes', this.bien);

    req.subscribe({
      next: (response: any) => {
        this.saving = false;
        this.editing = false;
        this.bienOriginal = null;
        terminar?.();
        if (response?.pendienteAutorizacion) {
          this.notificar('pendiente', 'Solicitud enviada',
            'Tu solicitud #' + response.idSolicitud + (esNuevo ? ' de alta' : ' de cambio') + ' espera la autorización del administrador.');
        } else {
          this.notificar('exito', esNuevo ? '¡Bien creado!' : 'Bien modificado',
            esNuevo ? 'El bien ya forma parte del inventario. Su formato de alta quedó guardado en Documentos.' : 'Los cambios se guardaron correctamente.',
            3500, esNuevo && this.isAdmin
              ? { texto: '📄 Ver formato de alta', fn: () => this.verOFormato(response?.idDocumento, () => this.formatoDesdeDatos('ALTA', formulario, {
                    folio: this.folio('ALTA', null, response?.id_bien ?? formulario.clave), solicito: this.miNombre, autorizo: this.miNombre,
                    observaciones: 'Alta registrada directamente por el administrador.' })) }
              : undefined);
        }
        this.loadBienes();
        if (response?.pendienteAutorizacion) this.loadSolicitudes();
      },
      error: e => {
        this.saving = false;
        const error = this.mensajeError(e, 'No se pudo guardar el bien.');
        if (terminar) terminar(error);
        else this.notificar('error', 'No se pudo guardar', error, 0);
      }
    });
  }

  deleteBien(id: number): void {
    const bien = this.bienes.find(b => b.id_bien === id) || this.bienSeleccionado;
    this.abrirDialogo({
      tipo: 'peligro',
      icono: '⚠',
      titulo: this.isAdmin ? '¿Deseas dar de baja este bien?' : '¿Solicitar la baja de este bien?',
      mensaje: this.isAdmin
        ? 'Dejará de aparecer en el inventario activo. El registro se conserva en el historial.'
        : 'Se enviará una solicitud al administrador. El bien seguirá activo hasta que la autorice.',
      textoConfirmar: this.isAdmin ? 'Sí, dar de baja' : 'Enviar solicitud de baja',
      bien,
      accion: terminar => {
        this.http.delete<any>(this.api + '/bienes/' + id).subscribe({
          next: response => {
            terminar();
            this.bienSeleccionado = null;
            if (response?.pendienteAutorizacion) {
              this.notificar('pendiente', 'Solicitud de baja enviada', 'Tu solicitud #' + response.idSolicitud + ' espera la autorización del administrador.');
            } else {
              this.notificar('baja', 'Bien dado de baja', 'Ya no aparece en el inventario activo. Su formato de baja quedó guardado en Documentos.', 3500,
                bien ? { texto: '📄 Ver formato de baja', fn: () => this.verOFormato(response?.idDocumento, () => this.formatoDesdeBien('BAJA', bien, {
                  folio: this.folio('BAJA', null, bien.id_bien), solicito: this.miNombre, autorizo: this.miNombre,
                  observaciones: 'Baja registrada directamente por el administrador.' })) } : undefined);
            }
            this.loadBienes();
            if (response?.pendienteAutorizacion) this.loadSolicitudes();
          },
          error: e => terminar(this.mensajeError(e, 'No se pudo dar de baja el bien.'))
        });
      }
    });
  }

  statusClass(b: any): string {
    const status = String(b.estatus_inventario || b.estado_fisico || '').toUpperCase();

    if (status === 'BAJA' || status === 'MALO') return 'status-bad';
    if (status === 'EN_REPARACION' || status === 'REGULAR') return 'status-warning';
    if (status === 'ACTIVO' || status === 'BUENO') return 'status-good';

    return 'status-neutral';
  }

  displayStatus(b: any): string {
    const status = String(b.estatus_inventario || b.estado_fisico || 'SIN ESTADO').toUpperCase();
    return status.replace(/_/g, ' ');
  }

  // =========================
  // CATÁLOGOS
  // =========================

  loadCatalog(key: string = this.catalog): void {
    if (!this.logged) return;
    this.http.get<any[]>(this.api + '/catalogos/' + key).subscribe({
      next: v => {
        this.catalogs[key] = v || [];
        if (key === this.catalog) this.catalogError = '';
      },
      error: e => {
        this.catalogs[key] = [];
        if (key === this.catalog) this.catalogError = this.mensajeError(e, 'No fue posible consultar el catálogo.');
      }
    });
  }

  selectCatalog(key: string): void {
    this.catalog = key;
    this.catalogFilter = '';
    this.catalogShowInactive = false;
    this.catalogPage = 1;
    this.catalogError = '';
    this.loadCatalog(key);
  }

  // Valor legible de una columna de la tabla
  valorColumna(x: any, col: any): string {
    const v = x[col.key];
    if (col.bool) return v ? 'Sí' : 'No';
    return v === null || v === undefined || String(v).trim() === '' ? '—' : String(v);
  }

  // Opciones de un select: solo registros activos, más el valor ya asignado aunque esté inactivo
  opcionesCampo(f: any): { valor: any; texto: string }[] {
    const o = f.opciones;
    const actual = this.catForm[f.key];
    return (this.catalogs[o.catalogo] || [])
      .filter((x: any) => x.estatus !== false || String(x[o.valor]) === String(actual))
      .map((x: any) => ({
        valor: f.texto ? String(x[o.valor]) : x[o.valor],
        texto: x[o.texto] + (x.estatus === false ? ' (inactivo)' : '')
      }))
      .sort((a: any, b: any) => a.texto.localeCompare(b.texto));
  }

  newCatalogItem(): void {
    this.catalogEditId = null;
    this.catForm = {};
    this.catalogConfig.campos.forEach((f: any) => {
      this.catForm[f.key] = f.tipo === 'check' ? f.key === 'estatus' : (f.tipo === 'select' ? null : '');
    });
    this.catalogFormError = '';
    this.catalogFieldError = '';
    this.catalogModal = true;
  }

  editCatalogItem(x: any): void {
    this.catalogEditId = x[this.catalogConfig.id];
    this.catForm = {};
    this.catalogConfig.campos.forEach((f: any) => {
      const v = x[f.key];
      if (f.tipo === 'check') this.catForm[f.key] = v !== false;
      else if (f.tipo === 'select') this.catForm[f.key] = v === null || v === undefined || v === '' ? null : (f.texto ? String(v) : v);
      else this.catForm[f.key] = v ?? '';
    });
    // Un valor de texto que no existe en el catálogo (ej. entidad inválida) queda sin seleccionar
    this.catalogConfig.campos.filter((f: any) => f.tipo === 'select' && f.texto).forEach((f: any) => {
      const existe = (this.catalogs[f.opciones.catalogo] || []).some((o: any) => String(o[f.opciones.valor]) === this.catForm[f.key]);
      if (!existe) this.catForm[f.key] = null;
    });
    this.catalogFormError = '';
    this.catalogFieldError = '';
    this.catalogModal = true;
  }

  saveCatalog(): void {
    if (this.catalogSaving) return;
    const config = this.catalogConfig;

    const faltante = config.campos.find((f: any) => {
      const v = this.catForm[f.key];
      return f.requerido && (v === null || v === undefined || String(v).trim() === '');
    });
    if (faltante) {
      this.catalogFieldError = faltante.key;
      this.catalogFormError = 'Indica: ' + faltante.label.toLowerCase() + '.';
      return;
    }

    const datos: any = {};
    config.campos.forEach((f: any) => {
      const v = this.catForm[f.key];
      datos[f.key] = typeof v === 'string' ? (v.trim() === '' ? null : v.trim()) : v;
    });

    this.catalogSaving = true;
    this.catalogFormError = '';
    this.catalogFieldError = '';
    const url = this.api + '/catalogos/' + this.catalog;
    const esNuevoRegistro = !this.catalogEditId;
    const req = this.catalogEditId ? this.http.put<any>(url + '/' + this.catalogEditId, datos) : this.http.post<any>(url, datos);

    req.subscribe({
      next: r => {
        this.catalogSaving = false;
        this.catalogModal = false;
        this.notificar('exito', esNuevoRegistro ? '¡Registro creado!' : 'Registro modificado',
          esNuevoRegistro ? 'Se agregó al catálogo correctamente.' : 'Los cambios se guardaron correctamente.');
        this.refreshAfterCatalogChange();
      },
      error: e => {
        this.catalogSaving = false;
        this.catalogFormError = this.mensajeError(e, 'No se pudo guardar. Verifica los datos.');
      }
    });
  }

  deleteCatalogItem(x: any): void {
    const config = this.catalogConfig;
    const nombre = this.valorColumna(x, config.columnas[0]);
    this.abrirDialogo({
      tipo: 'peligro',
      icono: config.estatus ? '⏻' : '⚠',
      titulo: (config.estatus ? '¿Desactivar ' : '¿Eliminar ') + config.singular.toLowerCase() + '?',
      mensaje: config.estatus
        ? '"' + nombre + '" dejará de aparecer en las listas, pero se conserva el historial. Podrás reactivarlo después.'
        : '"' + nombre + '" se eliminará definitivamente. Esta acción no se puede deshacer.',
      textoConfirmar: config.estatus ? 'Sí, desactivar' : 'Sí, eliminar',
      accion: terminar => {
        this.http.delete<any>(this.api + '/catalogos/' + this.catalog + '/' + x[config.id]).subscribe({
          next: r => {
            terminar();
            this.notificar('baja', config.estatus ? 'Registro desactivado' : 'Registro eliminado',
              config.estatus ? 'Ya no aparecerá en las listas; puedes reactivarlo cuando quieras.' : 'Se eliminó definitivamente del catálogo.');
            this.refreshAfterCatalogChange();
          },
          error: e => terminar(this.mensajeError(e, 'No se pudo completar la operación.'))
        });
      }
    });
  }

  reactivateCatalogItem(x: any): void {
    this.http.put<any>(this.api + '/catalogos/' + this.catalog + '/' + x[this.catalogConfig.id], { estatus: true }).subscribe({
      next: () => {
        this.notificar('exito', 'Registro reactivado', 'Vuelve a estar disponible en las listas.');
        this.refreshAfterCatalogChange();
      },
      error: e => this.catalogError = this.mensajeError(e, 'No se pudo reactivar el registro.')
    });
  }

  // Recarga el catálogo editado y los que muestran sus nombres (ej. sedes → ubicaciones, personas)
  private refreshAfterCatalogChange(): void {
    this.catalogError = '';
    const dependientes: any = {
      entidades: ['sedes'], sedes: ['ubicaciones', 'personas'], 'tipos-ubicacion': ['ubicaciones'],
      departamentos: ['puestos'], puestos: ['personas'], clasificaciones: ['subcategorias']
    };
    [this.catalog, ...(dependientes[this.catalog] || [])].forEach(k => this.loadCatalog(k));
  }

  // =========================
  // IMPORTACIÓN DE EXCEL
  // =========================

  openExcel(): void {
    if (!this.isAdmin) {
      this.error = 'Solo el administrador puede importar archivos Excel.';
      return;
    }
    this.excelModal = true;
    this.excelFile = null;
    this.excelRows = [];
    this.excelHeaders = [];
    this.excelError = '';
    this.excelImportResult = '';
    this.excelImporting = false;
  }

  closeExcel(): void {
    if (this.excelImporting) return;

    this.excelModal = false;
    this.excelFile = null;
    this.excelRows = [];
    this.excelHeaders = [];
    this.excelError = '';
    this.excelImportResult = '';
  }

  async handleExcel(event: Event): Promise<void> {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];

    if (!file) return;

    this.excelError = '';
    this.excelImportResult = '';
    this.excelRows = [];
    this.excelHeaders = [];

    const extension = file.name.split('.').pop()?.toLowerCase();

    if (!['xlsx', 'xls'].includes(extension || '')) {
      this.excelError = 'Selecciona un archivo Excel .xlsx o .xls.';
      input.value = '';
      return;
    }

    this.excelFile = file;

    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array' });
      const firstSheet = workbook.SheetNames[0];

      if (!firstSheet) {
        throw new Error('El archivo no contiene hojas.');
      }

      const worksheet = workbook.Sheets[firstSheet];
      const rows = XLSX.utils.sheet_to_json<any>(worksheet, {
        defval: '',
        raw: false
      });

      console.log('CONTENIDO DEL EXCEL POR FILAS:', rows);

      if (!rows.length) {
        throw new Error('La hoja seleccionada no contiene registros.');
      }

      this.excelRows = rows;
      this.excelHeaders = Object.keys(rows[0]);

    } catch (e: any) {
      this.excelFile = null;
      this.excelRows = [];
      this.excelHeaders = [];
      this.excelError = e.message || 'No se pudo leer el archivo.';
      input.value = '';
    }
  }

  confirmExcel(): void {
    if (!this.excelFile || !this.excelRows.length || this.excelImporting) {
      return;
    }

    const token = localStorage.getItem('dozzer_token');

    if (!token) {
      this.excelError = 'Tu sesión ha expirado. Inicia sesión nuevamente.';
      return;
    }

    this.excelImporting = true;
    this.excelError = '';
    this.excelImportResult = '';
    this.aviso = null;

    const headers = new HttpHeaders({
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    });

    this.http.post<any>(
      this.api + '/bienes/importar',
      { rows: this.excelRows },
      { headers }
    ).subscribe({
      next: response => {
        this.excelImporting = false;

        const total = Number(response?.total ?? this.excelRows.length);
        const importados = Number(response?.importados ?? 0);
        const rechazados = Number(response?.rechazados ?? 0);

        this.excelImportResult =
          `Proceso terminado: ${importados} de ${total} registros importados.`;

        if (rechazados > 0) {
          const errores = Array.isArray(response?.errores)
            ? response.errores.map((e: any) =>
                typeof e === 'string'
                  ? e
                  : e.mensaje || e.error || JSON.stringify(e)
              )
            : [];

          this.excelError =
            `${rechazados} registro(s) no se pudieron importar:\n` +
            errores.slice(0, 10).join('\n');

          if (errores.length > 10) {
            this.excelError += `\nY ${errores.length - 10} error(es) más.`;
          }
        }

        if (importados > 0) {
          this.notificar('exito', 'Importación completa', `Se importaron ${importados} de ${total} bienes correctamente.`);
          this.loadBienes();
        }

        if (rechazados === 0 && importados > 0) {
          this.closeExcel();
        }
      },
      error: error => {
        this.excelImporting = false;

        this.excelError =
          error.error?.message ||
          error.error?.error ||
          `Error ${error.status}: No fue posible importar los registros.`;
      }
    });
  }
}
