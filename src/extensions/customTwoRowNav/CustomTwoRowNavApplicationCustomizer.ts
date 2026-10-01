import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';

const LOG_SOURCE: string = 'CustomTwoRowNavApplicationCustomizer';

export interface ICustomTwoRowNavApplicationCustomizerProperties {
  showBrandHeader?: boolean;
  hideDefaultSharePointNav?: boolean;
}

export default class CustomTwoRowNavApplicationCustomizer
  extends BaseApplicationCustomizer<ICustomTwoRowNavApplicationCustomizerProperties> {

  private _topPlaceholder: PlaceholderContent | undefined;
  private _outsideClickListener: ((event: MouseEvent) => void) | undefined;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initializing CustomTwoRowNav Application Customizer');

    // Wire up placeholder event handler to support SPA page transitions in SharePoint
    this.context.placeholderProvider.changedEvent.add(this, this._renderHeader);
    this._renderHeader();

    return Promise.resolve();
  }

  private _renderHeader(): void {
    // Check if the top placeholder is available
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this._onDispose.bind(this) }
      );
    }

    if (!this._topPlaceholder) {
      Log.warn(LOG_SOURCE, 'The expected placeholder (Top) was not found.');
      return;
    }

    if (this._topPlaceholder.domElement) {
      this._injectStyles();
      this._topPlaceholder.domElement.innerHTML = this._buildNavigationHtml();
      this._bindEvents();
    }
  }

  private _buildNavigationHtml(): string {
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    const showBrandHeader = this.properties.showBrandHeader ?? true;

    const brandHeaderHtml = showBrandHeader ? `
      <!-- TOP UTILITY / BRAND ROW -->
      <div class="sirva-brand-bar">
        <div class="sirva-brand-left">
          <a href="${siteUrl}" class="sirva-logo">
            SIRVA<sup>&reg;</sup>
          </a>
          <div class="sirva-brand-sep"></div>
          <div class="sirva-tagline">
            <span>People</span>
            <span>Move</span>
            <span>Possibilities</span>
          </div>
        </div>

        <div class="sirva-brand-right">
          <div class="sirva-search-pill">
            <svg class="sirva-search-icon" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
              <circle cx="11" cy="11" r="8"></circle>
              <line x1="21" y1="21" x2="16.65" y2="16.65"></line>
            </svg>
            <input type="text" placeholder="Search across Sirva..." class="sirva-search-input" />
          </div>

          <div class="sirva-util-icons">
            <button type="button" class="sirva-icon-btn" title="Notifications" aria-label="Notifications">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"></path>
                <path d="M13.73 21a2 2 0 0 1-3.46 0"></path>
              </svg>
            </button>
            <button type="button" class="sirva-icon-btn" title="App launcher" aria-label="App launcher">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor">
                <circle cx="5" cy="5" r="2"></circle>
                <circle cx="12" cy="5" r="2"></circle>
                <circle cx="19" cy="5" r="2"></circle>
                <circle cx="5" cy="12" r="2"></circle>
                <circle cx="12" cy="12" r="2"></circle>
                <circle cx="19" cy="12" r="2"></circle>
                <circle cx="5" cy="19" r="2"></circle>
                <circle cx="12" cy="19" r="2"></circle>
                <circle cx="19" cy="19" r="2"></circle>
              </svg>
            </button>
            <div class="sirva-avatar" title="${this.context.pageContext.user.displayName || 'User'}">
              ${this._getUserInitials()}
            </div>
          </div>
        </div>
      </div>
    ` : '';

    return `
      <div id="sirva-custom-header" class="sirva-nav-root">
        ${brandHeaderHtml}

        <!-- ROW 1: PRIMARY NAVIGATION -->
        <div class="sirva-row-primary">
          <a href="${siteUrl}" class="sirva-item active">Home</a>

          <div class="sirva-item-group">
            <button type="button" class="sirva-item">
              About
              <svg class="sirva-chevron" width="10" height="6" viewBox="0 0 10 6" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
                <path d="M1 1L5 5L9 1"/>
              </svg>
            </button>
          </div>

          <div class="sirva-item-group">
            <button type="button" class="sirva-item">
              Life @ Sirva
              <svg class="sirva-chevron" width="10" height="6" viewBox="0 0 10 6" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
                <path d="M1 1L5 5L9 1"/>
              </svg>
            </button>
          </div>

          <div class="sirva-item-group">
            <button type="button" class="sirva-item">
              My Country
              <svg class="sirva-chevron" width="10" height="6" viewBox="0 0 10 6" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
                <path d="M1 1L5 5L9 1"/>
              </svg>
            </button>
          </div>

          <div class="sirva-item-group">
            <button type="button" class="sirva-item">
              Departments
              <svg class="sirva-chevron" width="10" height="6" viewBox="0 0 10 6" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
                <path d="M1 1L5 5L9 1"/>
              </svg>
            </button>
          </div>

          <div class="sirva-item-group">
            <button type="button" class="sirva-item">
              Comms
              <svg class="sirva-chevron" width="10" height="6" viewBox="0 0 10 6" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
                <path d="M1 1L5 5L9 1"/>
              </svg>
            </button>
          </div>
        </div>

        <!-- ROW 2: SECONDARY NAVIGATION -->
        <div class="sirva-row-secondary">
          
          <!-- People (Active Subnav Tab with Mega Menu) -->
          <div class="sirva-subnav-group active" id="sirva-dropdown-people">
            <button type="button" class="sirva-pill-btn active-tab" aria-haspopup="true" aria-expanded="false">
              People
              <svg class="sirva-chevron" width="10" height="6" viewBox="0 0 10 6" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
                <path d="M1 1L5 5L9 1"/>
              </svg>
            </button>

            <!-- MEGA MENU DROPDOWN -->
            <div class="sirva-mega-dropdown" role="menu">
              <!-- Column 1 -->
              <div class="sirva-mega-column">
                <h4 class="sirva-col-title">Your People Team</h4>
                <div class="sirva-col-links">
                  <a href="#">HR Services</a>
                  <a href="#">Benefits</a>
                  <a href="#">Payroll</a>
                  <a href="#">Policies</a>
                  <a href="#">Wellbeing</a>
                </div>
              </div>

              <!-- Column 2 -->
              <div class="sirva-mega-column">
                <h4 class="sirva-col-title">Leadership</h4>
                <div class="sirva-col-links">
                  <a href="#">Leaders Hub</a>
                  <a href="#">Executive Messages</a>
                  <a href="#">Leadership Resources</a>
                  <a href="#">Our Values</a>
                </div>
              </div>

              <!-- Column 3 -->
              <div class="sirva-mega-column">
                <h4 class="sirva-col-title">Global People</h4>
                <div class="sirva-col-links">
                  <a href="#">My Country</a>
                  <a href="#">Diversity, Equity & Inclusion</a>
                  <a href="#">People Programs</a>
                  <a href="#">Recognition</a>
                </div>
              </div>
            </div>
          </div>

          <div class="sirva-subnav-group">
            <button type="button" class="sirva-pill-btn">
              Leaders
              <svg class="sirva-chevron" width="10" height="6" viewBox="0 0 10 6" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
                <path d="M1 1L5 5L9 1"/>
              </svg>
            </button>
          </div>

          <div class="sirva-subnav-group">
            <button type="button" class="sirva-pill-btn">
              Learning
              <svg class="sirva-chevron" width="10" height="6" viewBox="0 0 10 6" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
                <path d="M1 1L5 5L9 1"/>
              </svg>
            </button>
          </div>

          <div class="sirva-subnav-group">
            <button type="button" class="sirva-pill-btn">
              Careers
              <svg class="sirva-chevron" width="10" height="6" viewBox="0 0 10 6" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
                <path d="M1 1L5 5L9 1"/>
              </svg>
            </button>
          </div>

          <div class="sirva-subnav-group">
            <button type="button" class="sirva-pill-btn">
              News @ Sirva
              <svg class="sirva-chevron" width="10" height="6" viewBox="0 0 10 6" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
                <path d="M1 1L5 5L9 1"/>
              </svg>
            </button>
          </div>

        </div>
      </div>
    `;
  }

  private _getUserInitials(): string {
    const displayName = this.context.pageContext.user.displayName || 'JD';
    const parts = displayName.trim().split(' ');
    if (parts.length >= 2) {
      return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
    }
    return parts[0].substring(0, 2).toUpperCase();
  }

  private _bindEvents(): void {
    if (!this._topPlaceholder || !this._topPlaceholder.domElement) return;

    const root = this._topPlaceholder.domElement;
    const dropdownGroups = root.querySelectorAll<HTMLElement>('.sirva-subnav-group, .sirva-item-group');

    dropdownGroups.forEach((group) => {
      const btn = group.querySelector('button');
      if (btn) {
        btn.onclick = (e) => {
          e.stopPropagation();
          const isCurrentlyOpen = group.classList.contains('is-open');

          // Close all others
          dropdownGroups.forEach((g) => {
            g.classList.remove('is-open');
            const b = g.querySelector('button');
            if (b) b.setAttribute('aria-expanded', 'false');
          });

          // Toggle if has mega-dropdown
          if (!isCurrentlyOpen && group.querySelector('.sirva-mega-dropdown')) {
            group.classList.add('is-open');
            btn.setAttribute('aria-expanded', 'true');
          }
        };
      }
    });

    // Close on outside click
    if (!this._outsideClickListener) {
      this._outsideClickListener = () => {
        if (!this._topPlaceholder || !this._topPlaceholder.domElement) return;
        this._topPlaceholder.domElement.querySelectorAll('.is-open').forEach((el) => {
          el.classList.remove('is-open');
          const btn = el.querySelector('button');
          if (btn) btn.setAttribute('aria-expanded', 'false');
        });
      };
      document.addEventListener('click', this._outsideClickListener);
    }
  }

  private _injectStyles(): void {
    const styleId = 'sirva-two-row-nav-styles';
    let styleTag = document.getElementById(styleId) as HTMLStyleElement;

    if (!styleTag) {
      styleTag = document.createElement('style');
      styleTag.id = styleId;
      document.head.appendChild(styleTag);
    }

    const hideDefaultNavCss = this.properties.hideDefaultSharePointNav ? `
      /* Optionally hide default SharePoint site navigation */
      div[data-automationid="SiteHeader"] div[role="navigation"],
      div[data-automationid="HorizontalNav"],
      .ms-HorizontalNav {
        display: none !important;
      }
    ` : '';

    styleTag.textContent = `
      /* Root Header Container */
      #sirva-custom-header.sirva-nav-root {
        width: 100%;
        background-color: #ffffff;
        font-family: "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif;
        position: relative;
        z-index: 1000;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05);
        color: #112a4a;
      }

      /* 1. TOP BRAND / SUITE BAR */
      #sirva-custom-header .sirva-brand-bar {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 0 40px;
        height: 60px;
        border-bottom: 1px solid #e3e8ef;
        background-color: #ffffff;
      }

      #sirva-custom-header .sirva-brand-left {
        display: flex;
        align-items: center;
        gap: 16px;
      }

      #sirva-custom-header .sirva-logo {
        font-size: 24px;
        font-weight: 800;
        letter-spacing: 0.5px;
        color: #112a4a;
        text-decoration: none;
        display: flex;
        align-items: center;
      }

      #sirva-custom-header .sirva-logo sup {
        font-size: 10px;
        font-weight: 600;
        top: -0.6em;
        margin-left: 2px;
      }

      #sirva-custom-header .sirva-brand-sep {
        width: 1.5px;
        height: 28px;
        background-color: #ccd7e4;
      }

      #sirva-custom-header .sirva-tagline {
        font-size: 11.5px;
        line-height: 1.2;
        font-weight: 500;
        color: #314d6f;
        display: flex;
        flex-direction: column;
      }

      #sirva-custom-header .sirva-brand-right {
        display: flex;
        align-items: center;
        gap: 20px;
      }

      #sirva-custom-header .sirva-search-pill {
        position: relative;
        display: flex;
        align-items: center;
      }

      #sirva-custom-header .sirva-search-icon {
        position: absolute;
        left: 12px;
        color: #7b8e9f;
        pointer-events: none;
      }

      #sirva-custom-header .sirva-search-input {
        width: 300px;
        height: 36px;
        padding: 0 16px 0 36px;
        background: #f4f6f9;
        border: 1px solid #dce2e9;
        border-radius: 20px;
        font-size: 13px;
        font-family: inherit;
        color: #112a4a;
        outline: none;
        transition: all 0.2s ease;
      }

      #sirva-custom-header .sirva-search-input:focus {
        background: #ffffff;
        border-color: #0066b2;
        box-shadow: 0 0 0 3px rgba(0, 102, 178, 0.12);
        width: 340px;
      }

      #sirva-custom-header .sirva-util-icons {
        display: flex;
        align-items: center;
        gap: 14px;
      }

      #sirva-custom-header .sirva-icon-btn {
        background: transparent;
        border: none;
        color: #435b75;
        cursor: pointer;
        display: flex;
        align-items: center;
        justify-content: center;
        width: 34px;
        height: 34px;
        border-radius: 50%;
        transition: background 0.15s ease, color 0.15s ease;
      }

      #sirva-custom-header .sirva-icon-btn:hover {
        background: #edf2f7;
        color: #0066b2;
      }

      #sirva-custom-header .sirva-avatar {
        width: 36px;
        height: 36px;
        border-radius: 50%;
        background: #112a4a;
        color: #ffffff;
        font-weight: 600;
        font-size: 13px;
        display: flex;
        align-items: center;
        justify-content: center;
        cursor: pointer;
      }

      /* 2. ROW 1: PRIMARY NAVIGATION */
      #sirva-custom-header .sirva-row-primary {
        display: flex;
        align-items: center;
        padding: 0 40px;
        height: 48px;
        border-bottom: 1px solid #e3e8ef;
        gap: 36px;
        background-color: #ffffff;
      }

      #sirva-custom-header .sirva-item {
        display: inline-flex;
        align-items: center;
        gap: 7px;
        height: 100%;
        font-size: 15px;
        font-weight: 500;
        color: #112a4a;
        text-decoration: none;
        cursor: pointer;
        position: relative;
        background: transparent;
        border: none;
        font-family: inherit;
        padding: 0;
        transition: color 0.15s ease;
      }

      #sirva-custom-header .sirva-item:hover {
        color: #0066b2;
      }

      /* Active indicator underline for Home */
      #sirva-custom-header .sirva-item.active {
        font-weight: 600;
        color: #112a4a;
      }

      #sirva-custom-header .sirva-item.active::after {
        content: "";
        position: absolute;
        bottom: -1px;
        left: 0;
        right: 0;
        height: 3px;
        background-color: #0066b2;
        border-radius: 2px 2px 0 0;
      }

      #sirva-custom-header .sirva-chevron {
        transition: transform 0.2s ease;
      }

      #sirva-custom-header .sirva-item:hover .sirva-chevron,
      #sirva-custom-header .sirva-pill-btn:hover .sirva-chevron,
      #sirva-custom-header .is-open .sirva-chevron {
        transform: translateY(1.5px);
      }

      /* 3. ROW 2: SECONDARY NAVIGATION (SUB-NAV) */
      #sirva-custom-header .sirva-row-secondary {
        display: flex;
        align-items: center;
        padding: 4px 40px;
        height: 46px;
        background-color: #ffffff;
        border-bottom: 1px solid #e3e8ef;
        gap: 24px;
        position: relative;
      }

      #sirva-custom-header .sirva-subnav-group {
        position: relative;
        display: inline-flex;
        align-items: center;
      }

      #sirva-custom-header .sirva-pill-btn {
        display: inline-flex;
        align-items: center;
        gap: 7px;
        padding: 6px 14px;
        font-size: 14.5px;
        font-weight: 500;
        color: #112a4a;
        border-radius: 6px;
        border: none;
        background: transparent;
        font-family: inherit;
        cursor: pointer;
        transition: all 0.15s ease;
        white-space: nowrap;
      }

      #sirva-custom-header .sirva-pill-btn:hover {
        color: #0066b2;
        background-color: rgba(238, 244, 250, 0.7);
      }

      /* Active pill state (People) */
      #sirva-custom-header .sirva-subnav-group.active .sirva-pill-btn,
      #sirva-custom-header .sirva-subnav-group.is-open .sirva-pill-btn {
        background-color: #eef4fa;
        color: #112a4a;
        font-weight: 600;
      }

      /* 4. MEGA MENU DROPDOWN PANEL */
      #sirva-custom-header .sirva-mega-dropdown {
        opacity: 0;
        visibility: hidden;
        transform: translateY(8px);
        transition: opacity 0.2s ease, transform 0.2s ease, visibility 0.2s;
        position: absolute;
        top: calc(100% + 4px);
        left: 0;
        background: #ffffff;
        min-width: 660px;
        border: 1px solid #dce3ed;
        border-top: none;
        border-radius: 0 0 8px 8px;
        box-shadow: 0 14px 32px rgba(17, 42, 74, 0.14), 0 3px 8px rgba(17, 42, 74, 0.04);
        padding: 26px 32px 32px 32px;
        display: grid;
        grid-template-columns: repeat(3, minmax(170px, 1fr));
        gap: 32px;
        z-index: 9999;
      }

      #sirva-custom-header .sirva-subnav-group:hover .sirva-mega-dropdown,
      #sirva-custom-header .sirva-subnav-group.is-open .sirva-mega-dropdown {
        opacity: 1;
        visibility: visible;
        transform: translateY(0);
      }

      #sirva-custom-header .sirva-mega-column {
        display: flex;
        flex-direction: column;
      }

      #sirva-custom-header .sirva-col-title {
        font-size: 14.5px;
        font-weight: 700;
        color: #112a4a;
        margin: 0 0 14px 0;
        padding: 0;
      }

      #sirva-custom-header .sirva-col-links {
        display: flex;
        flex-direction: column;
        gap: 10px;
      }

      #sirva-custom-header .sirva-col-links a {
        font-size: 13.5px;
        color: #38516d;
        text-decoration: none;
        font-weight: 400;
        transition: all 0.15s ease;
        line-height: 1.3;
      }

      #sirva-custom-header .sirva-col-links a:hover {
        color: #0066b2;
        transform: translateX(3px);
      }

      /* Responsive Media Queries */
      @media (max-width: 1024px) {
        #sirva-custom-header .sirva-brand-bar,
        #sirva-custom-header .sirva-row-primary,
        #sirva-custom-header .sirva-row-secondary {
          padding-left: 20px;
          padding-right: 20px;
        }
        #sirva-custom-header .sirva-row-primary {
          gap: 22px;
        }
        #sirva-custom-header .sirva-row-secondary {
          gap: 16px;
        }
      }

      @media (max-width: 768px) {
        #sirva-custom-header .sirva-search-pill {
          display: none;
        }
        #sirva-custom-header .sirva-row-primary,
        #sirva-custom-header .sirva-row-secondary {
          overflow-x: auto;
          white-space: nowrap;
          -webkit-overflow-scrolling: touch;
        }
        #sirva-custom-header .sirva-mega-dropdown {
          min-width: 88vw;
          grid-template-columns: 1fr;
          padding: 18px;
          gap: 18px;
        }
      }

      ${hideDefaultNavCss}
    `;
  }

  private _onDispose(): void {
    if (this._outsideClickListener) {
      document.removeEventListener('click', this._outsideClickListener);
      this._outsideClickListener = undefined;
    }
    Log.info(LOG_SOURCE, 'Disposed CustomTwoRowNav Application Customizer.');
  }
}
