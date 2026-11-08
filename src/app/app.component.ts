
import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { HttpClient, HttpHeaders } from '@angular/common/http';
import * as XLSX from 'xlsx';
import * as QRCode from 'qrcode';
import { API_URL, QR_FICHA_URL } from './config';

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

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, FormsModule],
  template: `
  <div class="app-shell">

    <!-- MENÚ LATERAL -->
    <aside class="sidebar" [class.collapsed]="menuCollapsed" *ngIf="logged && !mustChangePassword">
      <div class="brand">
        <div class="brand-icon">D</div>
        <div class="brand-text">
          <strong>DOZZER</strong>
          <small>Construcción y Asesoría S.A. de C.V.</small>
        </div>
      </div>

      <div class="menu-label">MENÚ PRINCIPAL</div>

      <button class="nav-item" [class.active]="section==='bienes'" (click)="section='bienes'; editing=false">
        <span class="nav-icon">▦</span>
        <span>Inventario</span>
      </button>

      <button class="nav-item" [class.active]="section==='solicitudes'" (click)="section='solicitudes'; loadSolicitudes()">
        <span class="nav-icon">▣</span>
        <span>{{ isAdmin ? 'Autorizaciones' : 'Solicitudes' }}</span>
        <span class="nav-badge" *ngIf="pendientesBadge > 0">{{ pendientesBadge }}</span>
      </button>

      <button class="nav-item" [class.active]="section==='documentos'" (click)="section='documentos'">
        <span class="nav-icon">▤</span>
        <span>Documentos</span>
      </button>

      <button class="nav-item" *ngIf="isAdmin" [class.active]="section==='catalogos'" (click)="section='catalogos'; loadCatalog()">
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

        <!-- Panel de marca -->
        <div class="lg-hero">
          <div class="lg-grid"></div>
          <div class="lg-glow"></div>

          <div class="lg-hero-top">
            <!-- LOGO: reemplazar este bloque por <img src="logo-dozzer.png" alt="DOZZER"> cuando llegue el logo -->
            <div class="lg-logo">D</div>
            <div>
              <strong class="lg-name">DOZZER</strong>
              <small class="lg-razon">Construcción y Asesoría S.A. de C.V.</small>
            </div>
          </div>

          <div class="lg-hero-body">
            <span class="lg-kicker">Control de activos</span>
            <h2>Todo tu inventario,<br>en un solo lugar.</h2>
            <p>Registra, ubica y da seguimiento a cada bien de la empresa, con autorizaciones y etiquetas QR.</p>

            <ul class="lg-points">
              <li>
                <span class="lg-point-ico"><svg viewBox="0 0 24 24"><path d="M4 4h7v7H4zM13 4h7v7h-7zM4 13h7v7H4zM13 13h3v3h-3zM20 13v.01M13 20h.01M17 20h3v-3"/></svg></span>
                <div><strong>Etiquetas QR</strong><small>Identifica cualquier bien con solo escanearlo.</small></div>
              </li>
              <li>
                <span class="lg-point-ico"><svg viewBox="0 0 24 24"><path d="M9 12l2 2 4-4"/><path d="M12 3l8 4v5c0 5-3.5 8-8 9-4.5-1-8-4-8-9V7z"/></svg></span>
                <div><strong>Cambios autorizados</strong><small>Cada alta, cambio o baja pasa por revisión.</small></div>
              </li>
              <li>
                <span class="lg-point-ico"><svg viewBox="0 0 24 24"><path d="M3 21h18M5 21V7l7-4 7 4v14M10 21v-4h4v4"/></svg></span>
                <div><strong>Todas las sedes</strong><small>Cada quien ve el inventario que le corresponde.</small></div>
              </li>
            </ul>
          </div>

          <small class="lg-hero-foot">© {{ anioActual }} DOZZER · Sistema de inventarios</small>
        </div>

        <!-- Formulario -->
        <div class="lg-side">
          <form class="lg-form" (ngSubmit)="doLogin()" autocomplete="on">
            <div class="lg-mobile-brand">
              <div class="lg-logo small">D</div>
              <strong>DOZZER</strong>
            </div>

            <h1>Bienvenido</h1>
            <p class="lg-sub">Inicia sesión con tu cuenta de DOZZER.</p>

            <label class="lg-label" for="lg-correo">Correo electrónico</label>
            <div class="lg-field" [class.invalid]="error">
              <svg viewBox="0 0 24 24"><rect x="3" y="5" width="18" height="14" rx="2"/><path d="M3 7l9 6 9-6"/></svg>
              <input id="lg-correo" name="correo" [(ngModel)]="login.correo" type="email" placeholder="tu.correo@dozzer.mx" autocomplete="username" (input)="error=''">
            </div>

            <label class="lg-label" for="lg-pass">Contraseña</label>
            <div class="lg-field" [class.invalid]="error">
              <svg viewBox="0 0 24 24"><rect x="4" y="11" width="16" height="10" rx="2"/><path d="M8 11V7a4 4 0 0 1 8 0v4"/></svg>
              <input id="lg-pass" name="password" [(ngModel)]="login.password" [type]="verPasswordLogin ? 'text' : 'password'" placeholder="Tu contraseña" autocomplete="current-password" (input)="error=''">
              <button type="button" class="lg-eye" (click)="verPasswordLogin = !verPasswordLogin" [title]="verPasswordLogin ? 'Ocultar contraseña' : 'Mostrar contraseña'">
                <svg *ngIf="!verPasswordLogin" viewBox="0 0 24 24"><path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12z"/><circle cx="12" cy="12" r="3"/></svg>
                <svg *ngIf="verPasswordLogin" viewBox="0 0 24 24"><path d="M3 3l18 18M10.6 5.1A10 10 0 0 1 12 5c6.5 0 10 7 10 7a17 17 0 0 1-3.2 4.1M6.6 6.6A17 17 0 0 0 2 12s3.5 7 10 7a9.6 9.6 0 0 0 5.4-1.6M9.9 9.9a3 3 0 0 0 4.2 4.2"/></svg>
              </button>
            </div>

            <div class="lg-error" *ngIf="error">
              <svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"/><path d="M12 8v5M12 16h.01"/></svg>
              <span>{{ error }}</span>
            </div>

            <button type="submit" class="lg-submit" [disabled]="loading || !login.correo || !login.password">
              <span class="mini-ring light" *ngIf="loading"></span>
              {{ loading ? 'Ingresando…' : 'Iniciar sesión' }}
              <svg *ngIf="!loading" viewBox="0 0 24 24"><path d="M5 12h14M13 6l6 6-6 6"/></svg>
            </button>

            <p class="lg-help">¿Primera vez? Entra con la contraseña temporal que te dio el área de inventarios; el sistema te pedirá crear la tuya.</p>
          </form>
        </div>
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
          <button class="menu-toggle" (click)="menuCollapsed=!menuCollapsed">☰</button>
          <div class="topbar-title">
            <small>DOZZER / {{ sectionLabel }}</small>
            <strong>{{ sectionTitle }}</strong>
          </div>
          <div class="topbar-right">
            <span class="online-dot"></span>
            <span class="online-label">Sistema activo</span>
            <div class="top-avatar">{{ initials }}</div>
          </div>
        </header>

        <div class="content-area">

          <!-- MENSAJES -->
          <div class="alert-success" *ngIf="message">
            ✓ {{ message }}
            <button (click)="message=''">×</button>
          </div>

          <!-- INVENTARIO -->
          <section *ngIf="section==='bienes'">

            <div class="inv-head">
              <div>
                <h1>Inventario</h1>
                <p class="inv-resumen" *ngIf="bienes.length; else resumenVacio">
                  <span><b>{{ bienes.length | number }}</b> bienes</span>
                  <span><b>{{ totalPiezas | number }}</b> piezas</span>
                  <span><b>{{ porcentaje(bienesBuenEstado) }}%</b> en buen estado</span>
                  <span><b>{{ bienesDisponibles | number }}</b> sin asignar</span>
                  <span class="inv-scope">{{ isAdmin ? 'Todas las sedes' : 'Tu sede' }}</span>
                </p>
                <ng-template #resumenVacio><p class="inv-resumen">{{ bienesCargando ? 'Cargando…' : 'Bienes y equipo de la empresa' }}</p></ng-template>
              </div>
              <div class="heading-actions">
                <div class="more-wrap">
                  <button class="btn-outline" (click)="menuMasInventario=!menuMasInventario">
                    Más opciones <span class="caret">▾</span>
                  </button>
                  <div class="picker-backdrop" *ngIf="menuMasInventario" (click)="menuMasInventario=false"></div>
                  <div class="more-menu" *ngIf="menuMasInventario">
                    <button [disabled]="!bienes.length" (click)="menuMasInventario=false; abrirModalEtiquetas()">
                      <span class="more-ico">⎙</span>
                      <div><strong>Imprimir etiquetas QR</strong><small>De los bienes filtrados o de la página</small></div>
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

                <!-- TABLA -->
                <div class="inventory-panel">

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
                          </div>

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
                </div>
              </div>
            </ng-container>

            <!-- IMPRESIÓN DE ETIQUETAS QR POR LOTE -->
            <div class="modal-backdrop" *ngIf="modalEtiquetas">
              <div class="modal-card large-modal">
                <div class="modal-header">
                  <div>
                    <h2>Imprimir etiquetas QR</h2>
                    <p>Genera una hoja lista para imprimir en papel carta o en hojas de etiquetas adhesivas.</p>
                  </div>
                  <button class="close-button" (click)="modalEtiquetas=false">×</button>
                </div>

                <div class="label-config">
                  <div class="label-options">
                    <h3>¿Qué bienes?</h3>
                    <label class="option-card" [class.selected]="etiquetasAlcance==='filtrados'">
                      <input type="radio" name="alcance" value="filtrados" [(ngModel)]="etiquetasAlcance" (ngModelChange)="generarQrMuestra()">
                      <div>
                        <strong>Bienes filtrados</strong>
                        <small>{{ bienesVista.length | number }} bienes{{ hayFiltrosBienes ? ' según tus filtros actuales' : ' (todo el inventario)' }}</small>
                      </div>
                    </label>
                    <label class="option-card" [class.selected]="etiquetasAlcance==='pagina'">
                      <input type="radio" name="alcance" value="pagina" [(ngModel)]="etiquetasAlcance" (ngModelChange)="generarQrMuestra()">
                      <div>
                        <strong>Página actual</strong>
                        <small>{{ paginaBienes.length }} bienes visibles en la tabla</small>
                      </div>
                    </label>

                    <h3>Tamaño de etiqueta</h3>
                    <label class="option-card" [class.selected]="etiquetasFormato==='estandar'">
                      <input type="radio" name="formato" value="estandar" [(ngModel)]="etiquetasFormato" (ngModelChange)="generarQrMuestra()">
                      <div>
                        <strong>Estándar · 10 por hoja</strong>
                        <small>2 columnas, 9.5 × 5 cm. Recomendada para mobiliario y equipo.</small>
                      </div>
                    </label>
                    <label class="option-card" [class.selected]="etiquetasFormato==='compacta'">
                      <input type="radio" name="formato" value="compacta" [(ngModel)]="etiquetasFormato" (ngModelChange)="generarQrMuestra()">
                      <div>
                        <strong>Compacta · 24 por hoja</strong>
                        <small>3 columnas, 6.4 × 3.3 cm. Para equipo pequeño o herramienta.</small>
                      </div>
                    </label>

                    <div class="warning-box" *ngIf="fichaQrDePrueba">
                      <strong>Dirección de prueba</strong>
                      <p>Las etiquetas abrirán la ficha desde {{ fichaQrHost }}, que solo funciona en la red de la oficina. No imprimas las etiquetas definitivas hasta configurar la dirección pública.</p>
                    </div>

                    <div class="warning-box" *ngIf="etiquetasSeleccion.length > 600">
                      <strong>Son muchas etiquetas</strong>
                      <p>Se generarán {{ etiquetasSeleccion.length | number }} etiquetas ({{ hojasEtiquetas | number }} hojas). Puede tardar unos segundos; considera filtrar por sede o clasificación.</p>
                    </div>
                  </div>

                  <div class="label-preview-wrap">
                    <h3>Vista previa</h3>
                    <div class="label-preview" [class.compacta]="etiquetasFormato==='compacta'" *ngIf="etiquetasSeleccion[0] as p">
                      <img class="lp-qr" *ngIf="qrMuestra" [src]="qrMuestra" alt="QR de muestra">
                      <div class="lp-info">
                        <div class="lp-brand"><span>D</span> DOZZER</div>
                        <div class="lp-clave">{{ p.clave }}</div>
                        <div class="lp-articulo">{{ p.articulo }}</div>
                        <div class="lp-meta" *ngIf="etiquetasFormato==='estandar'">{{ ubicacionEtiqueta(p) }}</div>
                        <div class="lp-id">ID {{ idCortoQr(p.codigo_qr) }}</div>
                      </div>
                    </div>
                    <p class="preview-note">{{ etiquetasSeleccion.length | number }} etiquetas · {{ hojasEtiquetas | number }} hoja(s) tamaño carta.
                      En el cuadro de impresión elige <b>escala 100%</b> y <b>sin márgenes</b> o "Predeterminado".<br>El QR guarda los datos actuales del bien; si después se editan, reimprime su etiqueta.</p>
                  </div>
                </div>

                <div class="modal-footer">
                  <button class="btn-outline" (click)="modalEtiquetas=false">Cancelar</button>
                  <button class="btn-primary" [disabled]="!etiquetasSeleccion.length || generandoEtiquetas" (click)="imprimirEtiquetas(etiquetasSeleccion)">
                    {{ generandoEtiquetas ? 'Generando...' : '⎙ Generar e imprimir ' + (etiquetasSeleccion.length | number) + ' etiquetas' }}
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

                <div class="drawer-section" *ngIf="b.caracteristicas">
                  <h3>Características</h3>
                  <div class="note-box">{{ b.caracteristicas }}</div>
                </div>

                <div class="drawer-section">
                  <h3>Etiqueta QR</h3>
                  <div class="qr-block">
                    <div class="qr-frame">
                      <img *ngIf="qrPreview" [src]="qrPreview" alt="Código QR del bien">
                      <span *ngIf="!qrPreview" class="mini-ring"></span>
                    </div>
                    <div class="qr-info">
                      <small>ID permanente</small>
                      <strong class="qr-id">{{ idCortoQr(b.codigo_qr) }}</strong>
                      <p>Al escanearlo se abre una ficha con los datos del bien, sin iniciar sesión. Si editas el bien, reimprime la etiqueta para actualizarlos.</p>
                      <button class="btn-outline btn-small" [disabled]="!b.codigo_qr" (click)="imprimirEtiquetas([b])">⎙ Imprimir etiqueta</button>
                    </div>
                  </div>
                </div>

                <div class="info-box" *ngIf="!isAdmin">
                  <strong>Los cambios requieren autorización</strong>
                  <p>Al editar o dar de baja este bien se enviará una solicitud al administrador.</p>
                </div>
              </div>

              <div class="drawer-footer">
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

                <div *ngIf="excelImportResult" style="margin: 0 24px 15px; color: #21835a; font-size: 12px;">
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
              <button class="stat-card stat-button" [class.stat-selected]="tabSolicitudes==='PENDIENTE'" (click)="tabSolicitudes='PENDIENTE'">
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
                  <span class="tab-count">{{ t.key==='TODAS' ? solicitudes.length : conteoSolicitudes(t.key) }}</span>
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
                    <tr *ngFor="let s of solicitudesFiltradas" class="clickable-row" (click)="abrirSolicitud(s)">
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
                      <td><span class="status" [ngClass]="claseEstadoSolicitud(s.estado)">{{ etiquetaEstado(s.estado) }}</span></td>
                      <td *ngIf="vistaSolicitudes==='mias'" class="resolution-cell">
                        <span *ngIf="s.estado==='PENDIENTE'" class="muted">En revisión</span>
                        <span *ngIf="s.estado!=='PENDIENTE'" [title]="s.motivoResolucion || ''">
                          {{ s.motivoResolucion || (s.estado==='APROBADA' ? 'Aprobada sin comentarios' : '—') }}
                        </span>
                      </td>
                      <td class="action-cell">
                        <button class="btn-outline btn-small" (click)="$event.stopPropagation(); abrirSolicitud(s)">
                          {{ vistaSolicitudes==='autorizacion' && s.estado==='PENDIENTE' && puedeResolver(s) ? 'Revisar' : 'Ver detalle' }}
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
                    <button class="btn-primary" (click)="cerrarSolicitud()">Cerrar</button>
                  </ng-template>
                </div>
              </div>
            </div>
          </section>

          <!-- DOCUMENTOS -->
          <section *ngIf="section==='documentos'">
            <div class="page-heading">
              <div>
                <div class="eyebrow">ARCHIVO DIGITAL</div>
                <h1>Documentos</h1>
                <p>Consulta y administra los documentos asociados a los bienes.</p>
              </div>
            </div>
            <div class="inventory-panel placeholder-panel">
              <div class="placeholder-icon">▤</div>
              <h2>Gestión documental</h2>
              <p>Aquí estarán disponibles las facturas, fotografías, resguardos y archivos PDF de cada bien.</p>
              <span class="coming-soon">Módulo en preparación</span>
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
    </main>
  </div>
  `,
  styles: [`
    :host{display:block;min-height:100vh;font-family:Inter,Arial,sans-serif;color:#243247}
    *{box-sizing:border-box}
    .app-shell{min-height:100vh;background:#f4f6fa;display:flex}
    .sidebar{width:260px;background:#17263d;color:#fff;display:flex;flex-direction:column;padding:25px 17px;flex-shrink:0;transition:.25s}
    .sidebar.collapsed{width:82px}
    .sidebar.collapsed .brand-text,.sidebar.collapsed .menu-label,.sidebar.collapsed .nav-item span:not(.nav-icon),.sidebar.collapsed .user-info{display:none}
    .brand{display:flex;align-items:center;gap:12px;margin-bottom:37px}
    .brand-icon{width:39px;height:39px;background:#fff;color:#193452;border-radius:11px;display:grid;place-items:center;font-size:23px;font-weight:900}
    .brand-text{display:flex;flex-direction:column;gap:4px;min-width:0}
    .brand-text strong{font-size:18px;letter-spacing:1.2px}
    .brand-text small{font-size:10px;color:#b9c6d7;white-space:nowrap}
    .menu-label{font-size:10px;letter-spacing:1.4px;color:#8192aa;font-weight:700;margin:5px 10px 12px}
    .nav-item{width:100%;border:0;background:transparent;color:#c1ccdb;display:flex;align-items:center;gap:14px;padding:13px 14px;border-radius:9px;text-align:left;margin-bottom:6px;font-size:13px;cursor:pointer}
    .nav-item:hover,.nav-item.active{background:#2b405e;color:#fff}
    .nav-item.active{box-shadow:inset 3px 0 #5e9be9}
    .nav-icon{font-size:19px;width:22px;text-align:center}
    .sidebar-bottom{margin-top:auto;border-top:1px solid #35455d;padding-top:18px;display:flex;align-items:center;gap:10px}
    .user-avatar,.top-avatar{width:36px;height:36px;border-radius:50%;background:#dce8f7;color:#244568;display:grid;place-items:center;font-size:12px;font-weight:800;flex-shrink:0}
    .user-info{display:flex;flex-direction:column;gap:4px;min-width:0;flex:1}
    .user-info strong{font-size:11px;overflow:hidden;text-overflow:ellipsis}
    .user-info small{font-size:10px;color:#9eacc0}
    .logout{background:transparent;color:#cbd5e1;border:0;font-size:20px;cursor:pointer}
    .main-content{flex:1;min-width:0}
    .topbar{height:76px;background:#fff;border-bottom:1px solid #e8edf3;display:flex;align-items:center;padding:0 32px;gap:18px}
    .menu-toggle{background:transparent;border:0;font-size:20px;color:#607087;cursor:pointer}
    .topbar-title{display:flex;flex-direction:column;gap:5px}
    .topbar-title small{font-size:10px;color:#8793a5;letter-spacing:1px}
    .topbar-title strong{font-size:14px;color:#25374e}
    .topbar-right{margin-left:auto;display:flex;align-items:center;gap:9px}
    .online-dot{width:8px;height:8px;border-radius:50%;background:#25a878}
    .online-label{font-size:11px;color:#6b7b90}
    .content-area{padding:30px 34px;max-width:1700px;margin:auto}
    .page-heading{display:flex;justify-content:space-between;align-items:center;gap:20px;margin-bottom:25px}
    .eyebrow{font-size:10px;font-weight:800;color:#6482a6;letter-spacing:1.4px;margin-bottom:8px}
    h1{font-size:27px;margin:0 0 7px;font-weight:750;letter-spacing:-.7px;color:#1c2d43}
    .page-heading p{margin:0;color:#8491a3;font-size:13px}
    .heading-actions{display:flex;gap:10px}
    button{font-family:inherit}
    .btn-primary,.btn-outline{border-radius:8px;padding:11px 16px;font-size:12px;font-weight:700;cursor:pointer;display:inline-flex;align-items:center;justify-content:center;gap:8px;transition:.2s}
    .btn-primary{background:#244c78;color:#fff;border:1px solid #244c78}
    .btn-primary:hover{background:#193b63}
    .btn-primary:disabled{opacity:.55;cursor:not-allowed}
    .btn-outline{background:#fff;color:#344c68;border:1px solid #dce3ec}
    .btn-outline:hover{background:#f4f7fb;border-color:#b7c6d9}
    .full{width:100%;margin-top:20px}
    .stats-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:15px;margin-bottom:23px}
    .stat-card{background:#fff;border:1px solid #e8edf3;border-radius:11px;padding:17px 19px;box-shadow:0 2px 7px #253b5507}
    .stat-top{display:flex;align-items:center;justify-content:space-between;color:#718096;font-size:12px}
    .stat-card>strong{display:block;font-size:27px;color:#1c2e45;margin:12px 0 4px;letter-spacing:-.8px}
    .stat-card>small{color:#9aa5b4;font-size:10px}
    .stat-icon{width:31px;height:31px;border-radius:8px;display:grid;place-items:center;font-weight:800}
    .blue{background:#e9f1fc;color:#4677b5}.purple{background:#f0edfc;color:#8068bd}.green{background:#e7f6ef;color:#26936a}.orange{background:#fff3e6;color:#c88a39}
    .inventory-panel{background:#fff;border:1px solid #e7ecf2;border-radius:12px;overflow:hidden;box-shadow:0 3px 12px #233c5708}
    .inventory-toolbar{padding:21px 23px 15px;display:flex;justify-content:space-between;align-items:center;gap:15px}
    .inventory-toolbar h2{font-size:15px;margin:0 0 5px;color:#25364c}
    .inventory-toolbar p{font-size:11px;color:#929eae;margin:0}
    .table-actions select{width:auto;min-width:155px}
    .search-row{padding:0 23px 17px;display:flex;align-items:center;gap:15px}
    .search-box{display:flex;align-items:center;gap:10px;border:1px solid #e1e7ef;border-radius:8px;padding:0 12px;height:39px;max-width:490px;flex:1}
    .search-box span{font-size:22px;color:#8493a6}
    .search-box input{border:0!important;outline:0!important;padding:0!important;box-shadow:none!important}
    input,select,textarea{width:100%;border:1px solid #dce3ec;border-radius:7px;padding:10px 11px;font-size:12px;color:#34465d;background:#fff;outline:none;font-family:inherit}
    input:focus,select:focus,textarea:focus{border-color:#7297c3;box-shadow:0 0 0 3px #7297c31c}
    .results-count{font-size:11px;color:#8491a2;white-space:nowrap}
    .table-container{width:100%;overflow:auto}
    table{width:100%;border-collapse:collapse;white-space:nowrap;text-align:left}
    thead{background:#f7f9fc}
    th{font-size:9px;letter-spacing:.65px;color:#8491a2;font-weight:800;padding:13px 15px;border-top:1px solid #edf0f4;border-bottom:1px solid #edf0f4}
    td{font-size:11px;color:#526176;padding:13px 15px;border-bottom:1px solid #eff2f6}
    tbody tr:hover{background:#fafbfd}
    .asset-key{color:#365d89;font-size:11px}
    .asset-name{display:flex;flex-direction:column;gap:5px}
    .asset-name strong{font-size:11px;color:#2b3d53}
    .asset-name small{font-size:10px;color:#9ba6b5}
    .quantity{background:#f1f4f8;padding:5px 9px;border-radius:5px;font-weight:700;color:#42566e}
    .status{font-size:9px;font-weight:800;padding:6px 9px;border-radius:20px;display:inline-block}
    .status-good{background:#e8f6ee;color:#21835a}
    .status-warning{background:#fff3df;color:#b27b22}
    .status-bad{background:#fdeced;color:#bd5056}
    .status-neutral{background:#edf1f6;color:#63738a}
    .row-actions{display:flex;gap:5px}
    .row-actions button{width:27px;height:27px;border:1px solid #e5eaf0;background:#fff;border-radius:6px;color:#60738a;cursor:pointer}
    .row-actions button:hover{background:#eef4fa}
    .row-actions .delete-action{color:#c65d62}
    .table-footer{display:flex;justify-content:space-between;padding:15px 22px;color:#8995a5;font-size:10px}
    .empty-state{text-align:center!important;padding:42px!important;color:#8491a2!important}
    .empty-state p{font-size:11px}
    .empty-icon{font-size:28px;color:#a8b6c7;margin-bottom:10px}
    .alert-success{background:#e8f6ee;color:#21835a;border:1px solid #cdebd9;border-radius:8px;padding:12px 15px;margin-bottom:18px;font-size:12px}
    .alert-success button{float:right;background:transparent;border:0;color:inherit;cursor:pointer}
    .modal-backdrop{position:fixed;inset:0;background:#15243b88;z-index:100;display:flex;align-items:center;justify-content:center;padding:20px}
    .modal-card{width:100%;max-width:590px;max-height:90vh;overflow:auto;background:#fff;border-radius:14px;box-shadow:0 20px 70px #15243b35}
    .large-modal{max-width:850px}
    .excel-modal{max-width:760px}
    .modal-header{padding:22px 24px;border-bottom:1px solid #edf0f4;display:flex;align-items:flex-start;justify-content:space-between}
    .modal-header h2{margin:0 0 6px;font-size:18px;color:#25364c}
    .modal-header p{margin:0;color:#8b97a7;font-size:11px}
    .close-button{background:transparent;border:0;color:#8290a2;font-size:25px;cursor:pointer}
    .form-section{padding:22px 24px}
    .form-section h3,.catalog-form h2{font-size:13px;color:#30445d;margin:0 0 17px}
    .form-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px 17px;margin-bottom:17px}
    .form-grid label,.form-section label,.login-card label{display:block;font-size:10px;font-weight:700;color:#64748a;margin-bottom:7px}
    .form-section textarea{resize:vertical}
    .modal-footer{display:flex;justify-content:flex-end;gap:10px;padding:16px 24px;border-top:1px solid #edf0f4}
    .excel-instructions{display:flex;gap:13px;padding:20px 24px 12px}
    .excel-icon{width:39px;height:39px;background:#e7f5ed;color:#21865d;border-radius:9px;display:grid;place-items:center;font-size:22px}
    .excel-instructions strong{font-size:12px;color:#33465e}
    .excel-instructions p,.preview-heading p{font-size:10px;color:#8995a5;margin:5px 0 0}
    .upload-zone{margin:10px 24px 20px;border:1.5px dashed #cbd7e4;border-radius:11px;min-height:165px;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:8px;background:#fafcff;cursor:pointer}
    .upload-zone:hover{border-color:#638dbb;background:#f5f9ff}
    .upload-zone.has-file{border-color:#58a77d;background:#f6fcf8}
    .upload-zone input{display:none}
    .upload-symbol{font-size:26px;color:#6486ac}
    .upload-zone strong{font-size:12px;color:#354b65}
    .upload-zone small{font-size:10px;color:#91a0b1}
    .upload-button{padding:8px 13px;margin-top:4px;font-size:10px}
    .excel-error{margin:0 24px 15px;color:#bb454b;font-size:11px}
    .preview-section{padding:0 24px 20px}
    .preview-heading{display:flex;justify-content:space-between;align-items:center;margin-bottom:12px}
    .preview-heading strong{font-size:12px;color:#33465e}
    .preview-table{border:1px solid #e8edf3;border-radius:7px;max-height:190px}
    .preview-table th,.preview-table td{padding:9px 11px}
    .preview-note{display:block;color:#8b97a7;font-size:10px;margin-top:10px}
    .detail-list{padding:22px 24px;display:grid;grid-template-columns:1fr 1fr;gap:17px}
    .detail-list div{display:flex;flex-direction:column;gap:5px}
    .detail-list small{font-size:10px;color:#8996a7}
    .detail-list strong{font-size:12px;color:#34485f}
    .placeholder-panel{text-align:center;padding:75px 20px}
    .placeholder-icon{font-size:38px;color:#7c9cbd}
    .placeholder-panel h2{font-size:17px}
    .placeholder-panel p{font-size:12px;color:#8995a5}
    .coming-soon{display:inline-block;margin-top:12px;padding:7px 12px;border-radius:20px;background:#eef3f9;color:#54779d;font-size:10px}
    .catalog-tabs{display:flex;gap:5px;flex-wrap:wrap;padding:18px 22px;border-bottom:1px solid #edf0f4}
    .catalog-tabs button{border:0;background:#f3f6fa;padding:9px 13px;border-radius:7px;font-size:11px;color:#687b91;cursor:pointer}
    .catalog-tabs button.selected{background:#244c78;color:white}
    .catalog-form{padding:22px}
    .catalog-form .form-grid>div{max-width:460px}
    .catalog-form label{display:block;font-size:10px;font-weight:700;color:#64748a;margin:12px 0 6px}
    .login-screen{min-height:100vh;display:grid;place-items:center;background:linear-gradient(135deg,#17263d,#294b70)}
    .login-card{background:#fff;width:min(410px,calc(100% - 35px));padding:38px;border-radius:15px;box-shadow:0 20px 70px #0c1b2d55}
    .login-brand{margin-bottom:30px;color:#243c58}
    .login-card h1{font-size:25px}
    .login-card>p{font-size:12px;color:#8995a5;margin-bottom:25px}
    .login-card label{margin-top:16px}
    .login-card input{height:42px}
    .error-message{color:#c24b52;font-size:11px;margin-top:12px}
    /* ===== INVENTARIO: CARGA ===== */
    /* ---------- INVENTARIO (vista limpia) ---------- */
    .inv-head{display:flex;justify-content:space-between;align-items:flex-end;gap:20px;margin-bottom:20px}
    .inv-head h1{font-size:26px;letter-spacing:-.6px;color:#14243a;margin:0 0 8px}
    .inv-resumen{display:flex;flex-wrap:wrap;align-items:center;gap:6px 18px;margin:0;font-size:12.5px;color:#7d8a9b}
    .inv-resumen b{color:#25364c;font-weight:700}
    .inv-resumen span{position:relative}
    .inv-resumen span+span::before{content:'';position:absolute;left:-10px;top:50%;width:3px;height:3px;border-radius:50%;background:#c3ccd8}
    .inv-scope{font-size:10.5px;font-weight:700;color:#4f6f95;background:#eef3f9;padding:3px 9px;border-radius:20px}
    .inv-resumen span.inv-scope::before{display:none}
    .more-wrap{position:relative}
    .caret{font-size:10px;color:#8a97a8}
    .more-menu{position:absolute;right:0;top:calc(100% + 6px);z-index:40;width:280px;background:#fff;border:1px solid #e3e9f1;border-radius:12px;box-shadow:0 14px 40px #1b2c4426;padding:6px;animation:fadeUp .15s ease both}
    .more-menu button{width:100%;display:flex;gap:12px;align-items:center;text-align:left;border:0;background:transparent;padding:10px;border-radius:9px;cursor:pointer;font-family:inherit}
    .more-menu button:hover:not(:disabled){background:#f3f6fa}
    .more-menu button:disabled{opacity:.45;cursor:not-allowed}
    .more-ico{width:34px;height:34px;border-radius:9px;background:#eef3f9;color:#244c78;display:grid;place-items:center;font-size:15px;flex-shrink:0}
    .more-menu strong{display:block;font-size:12.5px;color:#25364c}
    .more-menu small{display:block;font-size:11px;color:#8d99a9;margin-top:2px}
    .class-tabs{display:flex;gap:4px;padding:10px 16px 0;border-bottom:1px solid #edf1f6;overflow-x:auto;scrollbar-width:thin}
    .class-tabs button{flex-shrink:0;display:inline-flex;align-items:center;gap:7px;border:0;background:transparent;font-family:inherit;font-size:12px;font-weight:600;color:#6f7d8f;padding:10px 12px 12px;border-bottom:2px solid transparent;margin-bottom:-1px;cursor:pointer;white-space:nowrap;transition:color .15s,border-color .15s}
    .class-tabs button:hover{color:#25364c}
    .class-tabs button.active{color:#1c3a5e;border-bottom-color:#244c78}
    .class-tabs i{width:7px;height:7px;border-radius:50%;display:inline-block}
    .class-tabs span{font-size:10.5px;font-weight:700;color:#8a97a8;background:#f1f4f8;padding:2px 7px;border-radius:10px}
    .class-tabs button.active span{background:#e5eef8;color:#244c78}
    .inv-toolbar{display:flex;align-items:center;gap:12px;padding:14px 16px}
    .inv-toolbar .search-box{flex:1;min-width:0}
    .inv-toolbar+.search-row{padding:0 16px 12px}
    @media(max-width:700px){.inv-head{flex-direction:column;align-items:stretch}.inv-head .heading-actions>*{flex:1}.inv-toolbar{flex-wrap:wrap}.more-menu{left:0;right:auto}}

    /* ---------- LOGIN ---------- */
    .lg-screen{min-height:100vh;display:grid;grid-template-columns:minmax(0,1.1fr) minmax(0,1fr);background:#f4f6fa}
    .lg-hero{position:relative;overflow:hidden;color:#fff;background:linear-gradient(150deg,#0f1c2e 0%,#17304f 55%,#244c78 100%);padding:44px 56px;display:flex;flex-direction:column;justify-content:space-between;gap:40px}
    .lg-grid{position:absolute;inset:0;background-image:linear-gradient(#ffffff0d 1px,transparent 1px),linear-gradient(90deg,#ffffff0d 1px,transparent 1px);background-size:44px 44px;mask-image:radial-gradient(ellipse at 30% 40%,#000 20%,transparent 75%);-webkit-mask-image:radial-gradient(ellipse at 30% 40%,#000 20%,transparent 75%);pointer-events:none}
    .lg-glow{position:absolute;width:520px;height:520px;right:-180px;bottom:-200px;border-radius:50%;background:radial-gradient(circle,#4f8ad055,transparent 65%);pointer-events:none}
    .lg-hero>*{position:relative}
    .lg-hero-top{display:flex;align-items:center;gap:14px;animation:fadeUp .5s ease both}
    .lg-logo{width:46px;height:46px;border-radius:13px;background:#fff;color:#17304f;display:grid;place-items:center;font-size:25px;font-weight:900;box-shadow:0 8px 24px #0006;flex-shrink:0}
    .lg-logo img{width:100%;height:100%;object-fit:contain;border-radius:inherit}
    .lg-logo.small{width:38px;height:38px;font-size:21px;border-radius:11px;background:#17304f;color:#fff;box-shadow:none}
    .lg-name{display:block;font-size:20px;letter-spacing:2px}
    .lg-razon{display:block;font-size:11px;color:#a9bad0;margin-top:3px}
    .lg-hero-body{max-width:470px;animation:fadeUp .6s .08s ease both}
    .lg-kicker{display:inline-block;font-size:10px;font-weight:800;letter-spacing:1.6px;text-transform:uppercase;color:#9cc3f0;background:#ffffff12;border:1px solid #ffffff1f;padding:6px 12px;border-radius:20px}
    .lg-hero-body h2{font-size:38px;line-height:1.12;letter-spacing:-.8px;margin:18px 0 14px;font-weight:800}
    .lg-hero-body>p{font-size:14px;line-height:1.6;color:#bccbdd;margin:0 0 30px}
    .lg-points{list-style:none;margin:0;padding:0;display:grid;gap:16px}
    .lg-points li{display:flex;gap:14px;align-items:flex-start}
    .lg-point-ico{width:38px;height:38px;border-radius:11px;background:#ffffff14;border:1px solid #ffffff1f;display:grid;place-items:center;flex-shrink:0}
    .lg-point-ico svg{width:18px;height:18px;fill:none;stroke:#cfe0f5;stroke-width:2;stroke-linecap:round;stroke-linejoin:round}
    .lg-points strong{display:block;font-size:13px;margin-bottom:3px}
    .lg-points small{font-size:12px;color:#a9bad0}
    .lg-hero-foot{font-size:11px;color:#8ea2bb}
    .lg-side{display:grid;place-items:center;padding:40px 24px;background:#fff}
    .lg-form{width:min(380px,100%);animation:fadeUp .5s .12s ease both}
    .lg-mobile-brand{display:none;align-items:center;gap:10px;margin-bottom:28px;color:#17304f;font-size:17px;letter-spacing:1.6px}
    .lg-form h1{font-size:28px;letter-spacing:-.5px;color:#14243a;margin:0 0 6px}
    .lg-sub{font-size:13px;color:#7a889a;margin:0 0 30px}
    .lg-label{display:block;font-size:11px;font-weight:700;color:#4a5a6e;margin:0 0 7px}
    .lg-field{position:relative;display:flex;align-items:center;margin-bottom:18px}
    .lg-field>svg{position:absolute;left:14px;width:17px;height:17px;fill:none;stroke:#94a2b4;stroke-width:2;stroke-linecap:round;stroke-linejoin:round;pointer-events:none;transition:stroke .2s}
    .lg-field input{width:100%;height:48px;padding:0 44px;border:1.5px solid #dfe5ee;border-radius:11px;background:#f8fafc;font-size:14px;color:#14243a;transition:border-color .2s,box-shadow .2s,background .2s}
    .lg-field input::placeholder{color:#a8b3c1}
    .lg-field input:focus{outline:none;border-color:#244c78;background:#fff;box-shadow:0 0 0 4px #244c781f}
    .lg-field:focus-within>svg{stroke:#244c78}
    .lg-field.invalid input{border-color:#e2a1a6}
    .lg-eye{position:absolute;right:6px;width:36px;height:36px;border:0;background:transparent;border-radius:8px;display:grid;place-items:center;cursor:pointer}
    .lg-eye:hover{background:#eef2f7}
    .lg-eye svg{width:18px;height:18px;fill:none;stroke:#7a889a;stroke-width:2;stroke-linecap:round;stroke-linejoin:round}
    .lg-error{display:flex;gap:9px;align-items:flex-start;background:#fdf0f1;color:#b4404a;border:1px solid #f5d3d6;border-radius:10px;padding:11px 13px;font-size:12px;line-height:1.45;margin:-4px 0 16px;animation:fadeUp .25s ease both}
    .lg-error svg{width:16px;height:16px;flex-shrink:0;fill:none;stroke:currentColor;stroke-width:2;stroke-linecap:round}
    .lg-submit{width:100%;height:50px;border:0;border-radius:11px;background:linear-gradient(135deg,#244c78,#17304f);color:#fff;font-size:14px;font-weight:700;display:flex;align-items:center;justify-content:center;gap:10px;cursor:pointer;box-shadow:0 10px 24px #17304f33;transition:transform .15s,box-shadow .2s,opacity .2s;margin-top:6px}
    .lg-submit:hover:not(:disabled){transform:translateY(-1px);box-shadow:0 14px 30px #17304f44}
    .lg-submit:disabled{opacity:.55;cursor:not-allowed;box-shadow:none}
    .lg-submit svg{width:17px;height:17px;fill:none;stroke:#fff;stroke-width:2.2;stroke-linecap:round;stroke-linejoin:round}
    .lg-help{font-size:11.5px;line-height:1.55;color:#8592a3;margin:22px 0 0;padding-top:18px;border-top:1px solid #edf1f6}
    @media (max-width:900px){
      .lg-screen{grid-template-columns:1fr}
      .lg-hero{display:none}
      .lg-mobile-brand{display:flex}
      .lg-side{min-height:100vh;align-items:start;padding-top:14vh}
    }
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
    .loading-banner strong{font-size:14px;color:#25364c}
    .loading-banner p{margin:5px 0 0;font-size:12px;color:#8491a3}
    .loader-ring{width:36px;height:36px;border-radius:50%;border:3px solid #dfe8f3;border-top-color:#244c78;animation:spin .9s linear infinite;flex-shrink:0}
    .mini-ring{display:inline-block;width:12px;height:12px;border-radius:50%;border:2px solid #dfe8f3;border-top-color:#244c78;animation:spin .9s linear infinite;vertical-align:-2px}
    .refreshing{font-size:11px;color:#6b7b90;display:inline-flex;align-items:center;gap:7px}

    /* ===== INVENTARIO: RESUMEN POR CLASIFICACIÓN ===== */
    .overview-panel{background:#fff;border:1px solid #e7ecf2;border-radius:12px;padding:20px 22px;margin-bottom:20px;box-shadow:0 3px 12px #233c5708}
    .overview-head{display:flex;justify-content:space-between;align-items:flex-start;gap:12px;margin-bottom:15px}
    .overview-head h2{font-size:15px;margin:0 0 5px;color:#25364c}
    .overview-head p{font-size:11px;color:#929eae;margin:0}
    .overview-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(170px,1fr));gap:10px}
    .overview-card{position:relative;text-align:left;font-family:inherit;border:1px solid #e8edf3;background:#fbfcfe;border-radius:10px;padding:13px 14px;cursor:pointer;display:flex;flex-direction:column;gap:6px;transition:.2s}
    .overview-card:hover{border-color:#c9d6e6;background:#fff;transform:translateY(-1px);box-shadow:0 4px 12px #233c570d}
    .overview-card.selected{border-color:#5e8fc7;background:#fff;box-shadow:0 0 0 3px #5e8fc71f}
    .overview-card.dimmed{opacity:.55}
    .overview-dot{position:absolute;top:15px;right:14px;width:8px;height:8px;border-radius:50%}
    .overview-name{font-size:10px;font-weight:800;letter-spacing:.4px;color:#6a7b90;padding-right:16px;text-transform:uppercase;line-height:1.35}
    .overview-card strong{font-size:22px;color:#1c2e45;letter-spacing:-.6px}
    .overview-bar{height:5px;background:#edf1f6;border-radius:4px;overflow:hidden}
    .overview-bar span{display:block;height:100%;border-radius:4px}
    .overview-card small{font-size:10px;color:#97a3b3}
    .link-button.inline{display:inline;margin:0;white-space:nowrap}

    /* ===== INVENTARIO: TABLA ===== */
    .inventory-filters{flex-wrap:wrap}
    .inventory-filters select{width:auto;min-width:160px;height:39px}
    .search-box{position:relative}
    .clear-search{background:transparent;border:0;color:#8a97a8;font-size:18px;cursor:pointer;padding:0 2px}
    .inventory-table th.sortable{cursor:pointer;user-select:none}
    .inventory-table th.sortable:hover{color:#2f4d70}
    .sort-icon{font-size:8px;margin-left:4px;color:#c2ccd8}
    .sort-icon.active{color:#244c78}
    .inventory-table .num{text-align:center}
    .chevron-col{width:28px;color:#b5c0cd;font-size:16px;text-align:center}
    .clickable-row:hover .chevron-col{color:#244c78}
    .row-selected{background:#f1f6fc!important}
    .page-size{width:auto;height:30px;padding:4px 8px;font-size:11px}
    .column-picker{position:relative}
    .picker-backdrop{position:fixed;inset:0;z-index:40}
    .picker-menu{position:absolute;right:0;top:calc(100% + 6px);z-index:41;width:270px;background:#fff;border:1px solid #e1e7ef;border-radius:12px;box-shadow:0 14px 40px #15243b2e;animation:fadeUp .18s ease both}
    .picker-head{padding:14px 16px 10px;border-bottom:1px solid #eef1f5}
    .picker-head strong{display:block;font-size:12px;color:#25364c}
    .picker-head small{font-size:10px;color:#8b97a7}
    .picker-presets{display:flex;gap:6px;padding:10px 16px}
    .picker-presets button{flex:1;border:1px solid #dce3ec;background:#f7f9fc;border-radius:7px;padding:6px 4px;font-size:10px;font-weight:700;color:#4d6077;cursor:pointer}
    .picker-presets button:hover{background:#eaf1f9;border-color:#b7c6d9}
    .picker-list{max-height:290px;overflow:auto;padding:4px 8px 10px}
    .picker-list label{display:flex;align-items:center;gap:9px;padding:7px 8px;border-radius:6px;font-size:12px;color:#3d4f66;cursor:pointer}
    .picker-list label:hover{background:#f5f8fb}
    .picker-list label.locked{color:#97a3b3;cursor:default}
    .picker-list input{width:auto;margin:0}
    .picker-list small{margin-left:auto;font-size:9px;color:#aab5c3;text-transform:uppercase;letter-spacing:.5px}

    .filters-menu{width:320px;max-height:calc(100vh - 160px);display:flex;flex-direction:column;overflow:auto}
    .filters-active{border-color:#5e8fc7!important;color:#244c78!important;background:#f5f9fe!important}
    .filter-fields{display:grid;gap:10px;padding:12px 16px 14px}
    .filter-fields label{display:flex;flex-direction:column;gap:5px;font-size:10px;font-weight:700;color:#64748a}
    .filter-fields select{height:36px;font-weight:400}
    .picker-subhead{border-top:1px solid #eef1f5}
    .picker-footer{position:sticky;bottom:0;display:flex;justify-content:space-between;align-items:center;gap:10px;padding:12px 16px;border-top:1px solid #eef1f5;background:#fafbfd}
    .picker-footer .link-button:disabled{opacity:.4;cursor:default;text-decoration:none}
    .active-filters{display:flex;flex-wrap:wrap;gap:6px;align-items:center}
    .filter-chip{display:inline-flex;align-items:center;gap:5px;background:#eaf1f9;color:#2f5580;border:1px solid #d3e1f1;border-radius:20px;padding:5px 6px 5px 11px;font-size:11px;animation:fadeUp .15s ease both}
    .filter-chip b{font-weight:700}
    .filter-chip button{border:0;background:#d3e1f1;color:#2f5580;width:18px;height:18px;border-radius:50%;cursor:pointer;font-size:13px;line-height:1;display:grid;place-items:center;padding:0}
    .filter-chip button:hover{background:#244c78;color:#fff}

    /* ===== INVENTARIO: PANEL LATERAL ===== */
    .drawer-backdrop{position:fixed;inset:0;background:#15243b55;z-index:90;animation:fadeIn .2s ease both}
    .drawer{position:fixed;inset:0;margin:auto;width:min(780px,calc(100% - 32px));height:fit-content;max-height:calc(100vh - 48px);background:#fff;z-index:91;display:flex;flex-direction:column;border-radius:16px;overflow:hidden;box-shadow:0 24px 80px #15243b40;animation:zoomIn .22s ease both}
    .drawer-section dl{grid-template-columns:repeat(3,minmax(0,1fr))!important}
    .drawer-header{display:flex;justify-content:space-between;gap:12px;padding:24px 24px 18px;border-bottom:1px solid #edf0f4}
    .drawer-header h2{margin:6px 0 10px;font-size:19px;color:#1c2d43;line-height:1.3}
    .drawer-badges{display:flex;gap:6px;flex-wrap:wrap}
    .drawer-body{flex:1;overflow:auto;padding:6px 24px 20px}
    .drawer-section{padding:16px 0;border-bottom:1px solid #f0f3f7}
    .drawer-section:last-of-type{border-bottom:0}
    .drawer-section h3{font-size:10px;font-weight:800;letter-spacing:1px;color:#6482a6;margin:0 0 12px;text-transform:uppercase}
    .drawer-section dl{margin:0;display:grid;grid-template-columns:1fr 1fr;gap:14px 18px}
    .drawer-section dl div{min-width:0}
    .drawer-section dt{font-size:10px;color:#8996a7;margin-bottom:4px}
    .drawer-section dd{margin:0;font-size:12px;font-weight:600;color:#2f4359;overflow-wrap:anywhere}
    .drawer-footer{display:flex;gap:10px;justify-content:flex-end;padding:16px 24px;border-top:1px solid #edf0f4;background:#fafbfd}
    .drawer .info-box{margin-top:14px}
    @media(max-width:700px){.overview-grid{grid-template-columns:repeat(2,minmax(0,1fr))}.inventory-filters select{flex:1;min-width:0}.picker-menu{right:auto;left:0}.drawer{width:calc(100% - 20px);max-height:calc(100vh - 20px)}.drawer-section dl{grid-template-columns:1fr 1fr!important}.drawer-footer button{flex:1}.sk-row span:nth-child(n+4){display:none}}

    /* ===== DIÁLOGO DE CONFIRMACIÓN ===== */
    .confirm-backdrop{position:fixed;inset:0;background:#15243b80;z-index:200;display:flex;align-items:center;justify-content:center;padding:16px;animation:fadeIn .15s ease both;backdrop-filter:blur(2px)}
    .confirm-card{width:min(460px,100%);max-height:calc(100vh - 32px);overflow:auto;background:#fff;border-radius:18px;padding:28px 26px 22px;text-align:center;box-shadow:0 30px 90px #0b172845;animation:zoomIn .2s ease both}
    .confirm-icon{width:58px;height:58px;margin:0 auto 16px;border-radius:50%;display:grid;place-items:center;font-size:26px}
    .confirm-peligro .confirm-icon{background:#fdeced;color:#c0474e;box-shadow:0 0 0 8px #fdeced80}
    .confirm-confirmar .confirm-icon{background:#e9f1fc;color:#2f5f97;box-shadow:0 0 0 8px #e9f1fc80}
    .confirm-card h2{margin:0 0 8px;font-size:19px;color:#1c2d43;letter-spacing:-.3px}
    .confirm-card>p{margin:0 auto;max-width:360px;font-size:13px;color:#6b7b90;line-height:1.55}
    .confirm-asset{margin:18px 0 0;background:#f6f8fb;border:1px solid #e8edf3;border-radius:12px;padding:12px 14px;display:flex;flex-direction:column;gap:4px;align-items:center}
    .confirm-asset strong{font-size:13px;color:#2b3d53;text-transform:uppercase}
    .confirm-asset small{font-size:11px;color:#8b97a7}
    .confirm-changes{margin:16px 0 0;text-align:left;border:1px solid #e8edf3;border-radius:12px;overflow:hidden}
    .confirm-changes-title{font-size:10px;font-weight:800;letter-spacing:1px;color:#6482a6;text-transform:uppercase;padding:10px 14px;background:#f7f9fc;border-bottom:1px solid #edf0f4}
    .change-row{display:grid;grid-template-columns:110px minmax(0,1fr) 16px minmax(0,1fr);gap:8px;align-items:center;padding:9px 14px;border-bottom:1px solid #f0f3f7;font-size:11px}
    .change-row:last-child{border-bottom:0}
    .change-label{font-weight:700;color:#5b6c82}
    .change-old{color:#9aa6b6;text-decoration:line-through;overflow-wrap:anywhere}
    .change-arrow{color:#b5c0cd;text-align:center}
    .change-new{color:#8a5a12;font-weight:700;overflow-wrap:anywhere}
    .confirm-note{margin-top:14px;background:#fff8ea;border:1px solid #f3e1b8;color:#8a5a12;border-radius:10px;padding:10px 12px;font-size:11px;line-height:1.45;text-align:left}
    .confirm-card .inline-error{margin:14px 0 0;text-align:left}
    .confirm-actions{display:flex;gap:10px;margin-top:22px}
    .confirm-actions button{flex:1;padding:12px 14px}
    .btn-danger-solid{border-radius:8px;padding:11px 16px;font-size:12px;font-weight:700;cursor:pointer;display:inline-flex;align-items:center;justify-content:center;gap:8px;background:#c0474e;color:#fff;border:1px solid #c0474e;transition:.2s}
    .btn-danger-solid:hover{background:#a63a40}
    .btn-danger-solid:disabled{opacity:.6;cursor:not-allowed}
    .mini-ring.light{border-color:#ffffff55;border-top-color:#fff}
    @media(max-width:520px){.change-row{grid-template-columns:1fr;gap:2px}.change-arrow{display:none}.confirm-actions{flex-direction:column-reverse}}

    /* ===== ETIQUETAS QR ===== */
    .qr-block{display:flex;gap:16px;align-items:center}
    .qr-frame{width:118px;height:118px;flex-shrink:0;border:1px solid #e3e9f0;border-radius:10px;display:grid;place-items:center;background:#fff;padding:6px}
    .qr-frame img{width:100%;height:100%;image-rendering:pixelated}
    .qr-info{display:flex;flex-direction:column;align-items:flex-start;gap:4px;min-width:0}
    .qr-info small{font-size:10px;color:#8996a7}
    .qr-id{font-family:'Courier New',monospace;font-size:15px;letter-spacing:1.5px;color:#244c78}
    .qr-info p{font-size:11px;color:#7d8b9d;margin:2px 0 8px;line-height:1.45}
    .label-config{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1fr);gap:24px;padding:20px 24px}
    .label-config h3{font-size:10px;font-weight:800;letter-spacing:1px;color:#6482a6;margin:0 0 10px;text-transform:uppercase}
    .label-options h3:not(:first-child){margin-top:20px}
    .option-card{display:flex;gap:11px;align-items:flex-start;border:1px solid #e1e7ef;border-radius:10px;padding:12px 13px;margin-bottom:8px;cursor:pointer;transition:.15s}
    .option-card:hover{border-color:#b7c6d9}
    .option-card.selected{border-color:#5e8fc7;background:#f5f9fe;box-shadow:0 0 0 3px #5e8fc71a}
    .option-card input{width:auto;margin:2px 0 0}
    .option-card strong{display:block;font-size:12px;color:#25364c;margin-bottom:3px}
    .option-card small{font-size:11px;color:#8491a3;line-height:1.4}
    .label-options .warning-box{margin-top:12px}
    .label-preview-wrap{background:#f4f6fa;border-radius:12px;padding:16px;display:flex;flex-direction:column}
    .label-preview{background:#fff;border:1px solid #c9d3df;border-radius:9px;padding:14px;display:flex;gap:14px;align-items:center;box-shadow:0 6px 18px #233c5714;margin:auto 0;width:100%;max-width:340px;align-self:center}
    .label-preview.compacta{max-width:240px;padding:10px;gap:10px}
    .lp-qr{width:112px;height:112px;flex-shrink:0;image-rendering:pixelated}
    .label-preview.compacta .lp-qr{width:76px;height:76px}
    .lp-info{display:flex;flex-direction:column;gap:4px;min-width:0}
    .lp-brand{display:flex;align-items:center;gap:5px;font-size:9px;font-weight:800;letter-spacing:1px;color:#244c78}
    .lp-brand span{width:15px;height:15px;border-radius:4px;background:#17263d;color:#fff;display:grid;place-items:center;font-size:9px}
    .lp-clave{font-size:17px;font-weight:800;color:#17263d}
    .label-preview.compacta .lp-clave{font-size:13px}
    .lp-articulo{font-size:10px;font-weight:700;color:#33465e;text-transform:uppercase;line-height:1.25;overflow:hidden;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical}
    .lp-meta{font-size:9px;color:#5d6f86}
    .lp-id{font-family:'Courier New',monospace;font-size:9px;color:#5d6f86;letter-spacing:.8px}
    .label-preview-wrap .preview-note{margin-top:14px;line-height:1.5}
    @media(max-width:800px){.label-config{grid-template-columns:1fr}}

    /* ===== CATÁLOGOS ===== */
    .catalog-layout{display:grid;grid-template-columns:230px minmax(0,1fr);gap:18px;align-items:start}
    .catalog-nav{background:#fff;border:1px solid #e7ecf2;border-radius:12px;padding:12px;box-shadow:0 3px 12px #233c5708;position:sticky;top:16px}
    .catalog-group{font-size:9px;font-weight:800;letter-spacing:1.2px;color:#8a9ab0;margin:12px 8px 6px}
    .catalog-group:first-child{margin-top:4px}
    .catalog-nav button{width:100%;display:flex;align-items:center;justify-content:space-between;gap:8px;border:0;background:transparent;padding:9px 10px;border-radius:7px;font-size:12px;color:#4d6077;cursor:pointer;text-align:left}
    .catalog-nav button:hover{background:#f3f6fa}
    .catalog-nav button.selected{background:#eaf1f9;color:#244c78;font-weight:700}
    .catalog-nav button.selected .tab-count{background:#244c78;color:#fff}
    .catalog-panel{min-width:0}
    .catalog-main{color:#2b3d53;font-size:11px}
    .inactive-row td{opacity:.6}
    .inline-check{display:inline-flex!important;align-items:center;gap:8px;font-size:12px!important;font-weight:600!important;color:#4d6077!important;white-space:nowrap;cursor:pointer;margin:0!important}
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
    .password-form label{display:block;font-size:10px;font-weight:700;color:#64748a;margin:16px 0 7px}
    .password-form label:first-child{margin-top:0}
    .password-form input{height:42px}
    .password-rules{list-style:none;padding:0;margin:16px 0 0;display:grid;gap:6px}
    .password-rules li{font-size:11px;color:#97a3b3;display:flex;align-items:center;gap:8px}
    .password-rules li::before{content:'○';font-size:12px}
    .password-rules li.ok{color:#21835a}
    .password-rules li.ok::before{content:'✓';font-weight:800}
    .link-button{display:block;margin:14px auto 0;background:transparent;border:0;color:#5f7590;font-size:12px;cursor:pointer;text-decoration:underline}

    /* ===== BANDEJAS DE SOLICITUDES ===== */
    .nav-badge{margin-left:auto;background:#e8913a;color:#fff;font-size:10px;font-weight:800;min-width:20px;height:20px;padding:0 6px;border-radius:10px;display:inline-grid;place-items:center}
    .red{background:#fdeced;color:#bd5056}
    .stat-button{font-family:inherit;text-align:left;cursor:pointer;width:100%;transition:.2s}
    .stat-button:hover{border-color:#c9d6e6;transform:translateY(-1px)}
    .stat-button.stat-selected{border-color:#5e8fc7;box-shadow:0 0 0 3px #5e8fc71f}
    .segmented{display:inline-flex;background:#eef2f7;border-radius:9px;padding:3px}
    .segmented button{border:0;background:transparent;padding:8px 13px;border-radius:7px;font-size:12px;font-weight:700;color:#647890;cursor:pointer}
    .segmented button.selected{background:#fff;color:#244c78;box-shadow:0 1px 4px #1f3b5a1f}
    .request-tabs{display:flex;gap:4px;padding:14px 22px 0;border-bottom:1px solid #edf0f4;flex-wrap:wrap}
    .request-tabs button{border:0;background:transparent;padding:11px 14px;font-size:12px;font-weight:700;color:#7a8a9e;cursor:pointer;border-bottom:2px solid transparent;margin-bottom:-1px;display:inline-flex;align-items:center;gap:8px}
    .request-tabs button:hover{color:#2f4d70}
    .request-tabs button.selected{color:#244c78;border-bottom-color:#244c78}
    .tab-count{background:#eef2f7;color:#5d7088;border-radius:10px;padding:2px 8px;font-size:10px}
    .request-tabs button.selected .tab-count{background:#244c78;color:#fff}
    .request-filters{padding-top:17px}
    .op-filter{width:auto;min-width:180px}
    .op-badge{font-size:10px;font-weight:800;padding:5px 9px;border-radius:6px;display:inline-block;letter-spacing:.3px}
    .op-alta{background:#e9f1fc;color:#3c6ead}
    .op-modificacion{background:#f0edfc;color:#7059b3}
    .op-baja{background:#fdeced;color:#bd5056}
    .person-cell{display:flex;align-items:center;gap:9px}
    .mini-avatar{width:28px;height:28px;border-radius:50%;background:#e4edf8;color:#2f5580;display:grid;place-items:center;font-size:10px;font-weight:800;flex-shrink:0}
    .clickable-row{cursor:pointer}
    .action-cell{text-align:right}
    .btn-small{padding:7px 12px;font-size:11px}
    .resolution-cell{max-width:260px;overflow:hidden;text-overflow:ellipsis}
    .muted{color:#9aa6b6;font-weight:400}
    .inline-error{margin:0 22px 14px;background:#fdeced;border:1px solid #f5cfd2;color:#b2444b;border-radius:8px;padding:10px 13px;font-size:12px}
    .form-section .inline-error{margin:12px 0 0}
    .modal-title-row{display:flex;align-items:center;gap:10px;margin-bottom:6px}
    .modal-title-row h2{margin:0!important}
    .request-summary{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:16px;padding:20px 24px;background:#f9fbfd;border-bottom:1px solid #edf0f4}
    .request-summary div{display:flex;flex-direction:column;gap:6px;align-items:flex-start;min-width:0}
    .request-summary small{font-size:10px;color:#8996a7;font-weight:700;letter-spacing:.3px}
    .request-summary strong{font-size:12px;color:#2f4359;overflow-wrap:anywhere}
    .form-section+.form-section{padding-top:0}
    .section-hint{font-size:11px;color:#8491a3;margin:-8px 0 12px}
    .compare-table{border:1px solid #e8edf3;border-radius:8px}
    .compare-table table{white-space:normal}
    .compare-table td{vertical-align:top}
    .field-label{font-weight:700;color:#5b6c82!important;width:200px}
    .old-value{color:#97a3b3!important}
    .changed-row{background:#fff9ec}
    .changed-row:hover{background:#fff4dc!important}
    .changed-row .old-value{text-decoration:line-through}
    .changed-row .new-value{color:#8a5a12!important;font-weight:700}
    .note-box{background:#f6f8fb;border:1px solid #e6ebf2;border-radius:8px;padding:12px 14px;font-size:12px;color:#40536b;white-space:pre-line}
    .warning-box,.info-box,.resolution-box{border-radius:10px;padding:14px 16px;font-size:12px}
    .warning-box{background:#fdf1f1;border:1px solid #f3d3d5;color:#9c3f45}
    .info-box{background:#f1f6fc;border:1px solid #d8e5f4;color:#36577e}
    .warning-box p,.info-box p,.resolution-box p{margin:6px 0 0;font-size:11px;opacity:.9}
    .resolution-box.approved{background:#eef8f2;border:1px solid #cdebd9;color:#1f7a53}
    .resolution-box.rejected{background:#fdf1f1;border:1px solid #f3d3d5;color:#a6434a}
    .resolution-reason{margin-top:12px;background:#ffffffb3;border-radius:7px;padding:10px 12px;color:#3d4e63;white-space:pre-line}
    .resolution-reason small{display:block;font-size:10px;font-weight:700;color:#8291a4;margin-bottom:4px}
    .decision-section{border-top:1px solid #edf0f4;padding-top:20px!important}
    .decision-section label{display:block;font-size:10px;font-weight:700;color:#64748a;margin-bottom:7px}
    .decision-section textarea{resize:vertical}
    .input-error{border-color:#d1666c!important;box-shadow:0 0 0 3px #d1666c1c!important}
    .field-footer{display:flex;justify-content:space-between;gap:10px;margin-top:6px}
    .field-error{font-size:11px;color:#bd4b52}
    .char-count{font-size:10px;color:#9aa6b6;margin-left:auto}
    .btn-success,.btn-danger{border-radius:8px;padding:11px 16px;font-size:12px;font-weight:700;cursor:pointer;display:inline-flex;align-items:center;gap:8px;transition:.2s}
    .btn-success{background:#21835a;color:#fff;border:1px solid #21835a}
    .btn-success:hover{background:#1a6d4a}
    .btn-danger{background:#fff;color:#bd4b52;border:1px solid #efc4c7}
    .btn-danger:hover{background:#fdf1f1;border-color:#e3a3a8}
    .btn-success:disabled,.btn-danger:disabled,.btn-outline:disabled{opacity:.55;cursor:not-allowed}
    @media(max-width:900px){.request-summary{grid-template-columns:repeat(2,minmax(0,1fr))}}
    @media(max-width:700px){.request-tabs{padding:10px 12px 0}.request-tabs button{padding:10px 9px}.request-filters{flex-wrap:wrap}.op-filter{flex:1}.request-summary{padding:16px 17px}.heading-actions{flex-wrap:wrap}.segmented{width:100%}.segmented button{flex:1}.field-label{width:auto}.modal-footer{flex-wrap:wrap}.modal-footer button{flex:1}}
    @media(max-width:1100px){.stats-grid{grid-template-columns:repeat(2,minmax(0,1fr))}.content-area{padding:24px 20px}}
    @media(max-width:700px){.sidebar{width:72px;padding:20px 10px}.brand-text,.menu-label,.nav-item span:not(.nav-icon),.user-info{display:none}.nav-item{justify-content:center;padding:13px 5px}.sidebar-bottom{justify-content:center}.logout{display:none}.topbar{padding:0 15px}.content-area{padding:20px 12px}.page-heading{align-items:flex-start;flex-direction:column}.heading-actions{width:100%}.heading-actions button{flex:1;padding:10px 8px}.stats-grid{gap:9px}.stat-card{padding:13px}.stat-card>strong{font-size:23px}.inventory-toolbar{padding:17px 14px;align-items:flex-start;flex-direction:column}.search-row{padding:0 14px 14px}.table-footer{padding:13px}.form-grid{grid-template-columns:1fr}.detail-list{grid-template-columns:1fr}.modal-backdrop{padding:9px}.modal-card{max-height:95vh}.modal-header,.form-section{padding:17px}.modal-footer{padding:13px 17px}.upload-zone{margin:10px 17px 18px}.preview-section{padding:0 17px 17px}}
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
    { clave: 'caracteristicas', claveActual: 'caracteristicas', etiqueta: 'Características' }
  ];

  section = 'bienes';
  menuCollapsed = false;

  error = '';
  message = '';
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
  filtroSede: number | null = null;
  filtroClasificacion: number | null = null;
  filtroEstadoFisico = '';
  filtroAsignacion = '';
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
  menuMasInventario = false;
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
        { key: 'correo', label: 'Correo' }, { key: 'nombre_puesto', label: 'Puesto' }, { key: 'nombre_sede', label: 'Sede' }
      ],
      campos: [
        { key: 'nombre_completo', label: 'Nombre completo', requerido: true, max: 180, ancho: 'completo' },
        { key: 'numero_empleado', label: 'Número de empleado', max: 20 },
        { key: 'correo', label: 'Correo electrónico', tipo: 'email', max: 180, placeholder: 'nombre@dozzer.mx' },
        { key: 'id_puesto', label: 'Puesto', tipo: 'select', opciones: { catalogo: 'puestos', valor: 'id_puesto', texto: 'nombre_puesto' } },
        { key: 'id_sede', label: 'Sede', tipo: 'select', requerido: true, opciones: { catalogo: 'sedes', valor: 'id_sede', texto: 'nombre_sede' } },
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

  constructor(private http: HttpClient) {
    if (this.logged && !this.mustChangePassword) {
      this.loadAll();
    }
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
        this.message = 'Tu contraseña se actualizó correctamente.';
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

  get initials(): string {
    return this.email.split('@')[0].split(/[._-]/).map(x => x.charAt(0)).join('').substring(0, 2).toUpperCase() || 'AD';
  }

  get sectionTitle(): string {
    return ({
      bienes: 'Inventario',
      documentos: 'Documentos',
      solicitudes: 'Solicitudes',
      catalogos: 'Configuración'
    } as any)[this.section];
  }

  get sectionLabel(): string {
    return ({
      bienes: 'INVENTARIO',
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
    if (this.filtroAsignacion) lista.push({ clave: 'asignacion', etiqueta: 'Asignación', valor: this.filtroAsignacion === 'ASIGNADO' ? 'Asignados' : 'Sin asignar' });
    return lista;
  }

  quitarFiltro(clave: string): void {
    if (clave === 'sede') this.filtroSede = null;
    if (clave === 'clasificacion') this.filtroClasificacion = null;
    if (clave === 'estatus') this.statusFilter = '';
    if (clave === 'fisico') this.filtroEstadoFisico = '';
    if (clave === 'asignacion') this.filtroAsignacion = '';
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
    this.filtroSolicitudes = '';
    this.filtroOperacion = '';
    this.section = 'bienes';
  }

  loadAll(): void {
    this.loadBienes();
    Object.keys(this.catalogs).forEach(k => this.loadCatalog(k));
    this.loadSolicitudes();
  }

  // =========================
  // BANDEJAS DE SOLICITUDES
  // =========================

  get esRevisor(): boolean {
    return this.isAdmin || this.canAuthorize;
  }

  get pendientesBadge(): number {
    return this.vistaSolicitudes === 'autorizacion' ? this.conteoSolicitudes('PENDIENTE') : 0;
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
    this.construirCamposSolicitud(s);
  }

  cerrarSolicitud(): void {
    if (this.resolviendo) return;
    this.solicitudSeleccionada = null;
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
    const accion = aprobar ? 'aprobar' : 'rechazar';

    this.http.post<any>(
      this.api + '/solicitudes-bienes/' + s.idSolicitud + '/' + accion,
      { motivoResolucion: motivo || null }
    ).subscribe({
      next: r => {
        this.resolviendo = '';
        this.solicitudSeleccionada = null;
        this.message = aprobar
          ? 'Solicitud #' + s.idSolicitud + ' aprobada. El cambio ya se aplicó al inventario.'
          : 'Solicitud #' + s.idSolicitud + ' rechazada. El solicitante verá el motivo en su bandeja.';
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
          _busqueda: [b.clave, b.articulo, b.modelo, b.marca, b.numero_serie, b.responsable, b.usuario_directo,
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
    if (reiniciarPagina) this.paginaActualBienes = 1;
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

  limpiarFiltrosBienes(): void {
    this.filter = '';
    this.statusFilter = '';
    this.filtroSede = null;
    this.filtroClasificacion = null;
    this.filtroEstadoFisico = '';
    this.filtroAsignacion = '';
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
    if (!b?.codigo_qr) return;
    QRCode.toDataURL(this.contenidoQr(b), { errorCorrectionLevel: 'M', margin: 1, width: 220, color: { dark: '#17263d' } })
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
      this.message = 'Tu navegador bloqueó la ventana de impresión. Permite las ventanas emergentes para DOZZER e inténtalo de nuevo.';
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
      this.message = 'No fue posible generar las etiquetas QR.';
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
      estatusAsignacion: 'SIN_ASIGNAR'
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
      caracteristicas: b.caracteristicas
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
      this.message = 'No hiciste cambios en el bien.';
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

    const req = this.bien.idBien
      ? this.http.put(this.api + '/bienes/' + this.bien.idBien, this.bien)
      : this.http.post(this.api + '/bienes', this.bien);

    req.subscribe({
      next: (response: any) => {
        this.saving = false;
        this.editing = false;
        this.bienOriginal = null;
        terminar?.();
        this.message = response?.pendienteAutorizacion
          ? 'Tu solicitud #' + response.idSolicitud + ' fue enviada y está pendiente de autorización. Puedes darle seguimiento en Solicitudes.'
          : 'Bien guardado correctamente.';
        this.loadBienes();
        if (response?.pendienteAutorizacion) this.loadSolicitudes();
      },
      error: e => {
        this.saving = false;
        const error = this.mensajeError(e, 'No se pudo guardar el bien.');
        if (terminar) terminar(error);
        else this.message = error;
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
            this.message = response?.pendienteAutorizacion
              ? 'Tu solicitud de baja #' + response.idSolicitud + ' fue enviada para autorización.'
              : 'Bien dado de baja correctamente.';
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
    const req = this.catalogEditId ? this.http.put<any>(url + '/' + this.catalogEditId, datos) : this.http.post<any>(url, datos);

    req.subscribe({
      next: r => {
        this.catalogSaving = false;
        this.catalogModal = false;
        this.message = r?.mensaje || 'Catálogo actualizado correctamente.';
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
            this.message = r?.mensaje || 'Registro actualizado.';
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
        this.message = 'Registro reactivado correctamente.';
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
    this.message = '';

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
          this.message =
            `Se importaron ${importados} de ${total} bienes correctamente.`;
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
