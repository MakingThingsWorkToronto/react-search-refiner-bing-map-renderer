
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Text, Environment, EnvironmentType, DisplayMode } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneChoiceGroup,
  IPropertyPaneChoiceGroupOption,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  IPropertyPanePage
} from '@microsoft/sp-webpart-base';

import * as strings from 'BingMapsSearchWebPartStrings';
import BingMaps from '../../components/BingMap/BingMap';
import IBingMapProps from '../../components/BingMap/IBingMapProps';
import IBingMapsPolygonSearchWebPartWebPartProps from './IBingMapsPolygonSearchWebPartWebPartProps';
import { ResultService, ISearchEvent} from '../../services/ResultService/ResultService';
import { ISearchResults, ISearchResult } from '../../models/ISearchResult';
import IResultService from '../../services/ResultService/IResultService';
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect';
import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/PropertyFieldNumber';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import BaseTemplateService from '../../services/TemplateService/BaseTemplateService';
import TemplateService from '../../services/TemplateService/TemplateService';
import MockTemplateService from '../../services/TemplateService/MockTemplateService';
import BingMap from '../../components/BingMap/BingMap';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import times = require('lodash/times');


export default class BingMapsPolygonSearchWebPartWebPart extends BaseClientSideWebPart<IBingMapsPolygonSearchWebPartWebPartProps> {

  private _searchResults : ISearchResults;
  private _resultService: IResultService;
  private _map: React.ReactElement<IBingMapProps>;
  private _component : BingMaps;
  private FIELDS: string = "Title,Description,Location";
  private _propertyFieldCodeEditor = null;
  private _propertyFieldCodeEditorLanguages = null;
  private _templateService: BaseTemplateService;
  private _mapTypeIds : IPropertyPaneDropdownOption[] = [
    {key:"aerial",text:"Aerial"},
    {key:"canvasDark",text:"Dark"},
    {key:"canvasLight",text:"Light"},
    {key:"birdseye",text:"Birdseye"},
    {key:"grayscale",text:"Grayscale"},
    {key:"ordnanceSurvey",text:"Ordnance Survey (UK Only)"},
    {key:"road",text:"Road"},
    {key:"streetside",text:"Street Side"}
  ];
  
  protected onInit(): Promise<void> {
    
    if (Environment.type === EnvironmentType.Local) {
      this._templateService = new MockTemplateService(this.context.pageContext.cultureInfo.currentUICultureName);
    } else {
        this._templateService = new TemplateService(this.context.spHttpClient, this.context.pageContext.cultureInfo.currentUICultureName);
    }

    this._resultService = new ResultService();
    this._resultService.registerRenderer(
            this.componentId, 
            'Bing Maps', 
            'MapPin', 
            (e) => this.onChangeHappened(e), this.properties.columns.map((i) => { return i.name; }) as string[]
    );

    return Promise.resolve();
    
  }

  private setMockResults() : void {

    let results : ISearchResult[] = [ 
        { 
          Shape: "1,w6_8xusi2JwkluHgv6zB8vvlHu9s9C12hiL7runFw1x3Hlmga5uv-Eqsq4Op6iyN_i9-Bvug5B651_Bvw8f6izzGx_87NxrynG3nu1Ej-p9B3u6wC7s_hDhw9W74lO72-Ukx4kEsi9nD7yxuDp2u1BiqzO-9y-D-9juM6yo0CwnxqGluzTvn-xNx6mtK0py5Dpyl8B7s-vDzg81C2glvEoryxHgj--Qg-q8H6yu5I917uFmxogNnvlxQs3ikNqqkrK6o9-Fw6q3Fmr6iLhgjxK9k04Lioy4JjsnsO9pi-Vk6klVpgnyJgr64Bgo6qDi_3cm_15DmxqzC-hj0CylzxL2364RvvmmG59qpCn0ihF-ywoL71rmKnvzd336nC41_ihB43nqV9tztMkpjzb47lyD_s7sOtl7rjB88p6S53tlHl9tlS9gzyQ3v3nBkk3hD7s09C7ik-HjvpxF5u17Bk0-yB_mzhBqrg6I2wkwD2s76Hq1urEqzzzF0wtgCmyonP2v75RkgmvHqxsmDv838Ct6l-H4zn6J1zxvK5zrlDlsmrCrpx6Jmt5tD_-09Emkg2EwzztCs3x5D69w3Xj0mkKwtnkH_6_lM-4uzPwmjmOrnr4Pq0jsP34t7pBmoslQht_3Xo4ymR_pymR883zI19k6Bjwz9B0y0Wx935ImwryPh1pxKo9oiKp-2pCy1jd5n1Y0wvzBkzl7DukuFonjxC4smyFz-hkJm3u7F8vvzK_tkpHpy17F880rG980pGk1iyR17jyHrn5mOry0rB776bwtlvCnqnR79k-B1vivHi5lxJlvygB6m5nB24rL8n79C66jqB84j_GnnpoF4u87Kik5sC1y6pEizqlL3u-8Ttr6gX2nxb1s15Boqk_Hyqm2Co2y8BrihrexqttNzvrtLq629Jki8_H3pzjK0g5kNn81vCg-thC73kwE1461Eqr44CnsshD1y8oEmy2lC7912R90_gKjyqkIwg4vLmktgR4pzjKzzrcgv01Dg5yW6r8nB5u7nIxiwpLj1mgC9-14Bg9o-B5jnLrgqoCsr91Y-lynEu3uvB5iwIpykgB243rOoh1lIp0iiU_hm2Splmc2r91C6z0elv_hB_vp8Dk7_6B66zmF2479D32zKzw5Xp_o7H6rnxFkqs-Cj4u_Bo0wsBssnV09gwMww55Dl9m9B2t3xB4q1tD5g41C5r7qB4i_bj8ioDxjusGk6qaotnuC8zoM_yyckhnqFlupyC2hkkIkxq5LiqngGpjmrHpxmf__9Os7q6Dn66yBhs2nBr5mItvgQs17_H-0j-Hkvm3Czl40FmitmEq0iiF7lh-sC3-l5N6tmzZ4qjmUsg90N_kmoTt86vIs3tjE2vyrDuj8mC615pB3z9H5qn2JxkqoC9ky3BipxgEr25wMyxroJgr5mQzvw4Ou442Uuzp8Gmn-uLh6u9OhxgvK_0olCu7Sg1skIs3n8Gxv35I7vh4Fs1msEorx8EskmrK0wpzI2llmHklinG2w-qJ77pmHj57rM-hp7To0qsL94-gWjg8jC788kGljs3Jpu83O84slVp_v-Ilm6sL55mlC39xgN5q4xiBt85kmB-t13MmvnuHwnr4I_hjkH4k3jE38_kFv7njzBo771JkkjjW44l8Z1ng6bqsx5Nri0rTp_zuPlh1tY92h_O5o0kjB0jrja80mxS0p1qL6o51Ss-urM82xNz-o54B8lqweoqnnX03yr0Bu32h2Byh71zBp4oxfy0_0Kn63sQ_i__con1oS-39gOryxyMli10Hj2o3kB-835Zh81v4BrgnuyByy2g4B5o871Bq8n5iC3k4unBpsrrgBk811kBhxzz4Bms4hnBn4l1nBhgip-B30luiC9nq89Bs__p3Bn73i3Bz_l-9Fi_0_nBk9hrgEvgy7lChqt1sCgzupSjjk7E0mgua9mgx4Bu31nsD05-5kB5khw1B7vo1VzozmanmxjrBzgmtTjn8xiBwu7kkBgp5kf9iwrC16hyR39w3T_uvrP9pwooBrvmikBlys0gB3-g-Ig159D9wwhQ0vqnPy4tjOk_huX0lm0Y4k1gaool1nCzr23hDyt7ijGrnsywBi-h8Jqyek0w1T0qgsnB6ph9gC1z_z4Bx3mo6B15hi1Bj9su2D71yrvC016viE5vj6rDsnp4jEwzl_sC4lvkvBww-v2Bjo90kD755-kComlm_Cgsi8Njz-pCo1kiKiuplB9s7Im8_FssrzB8kxbjs4D6l6G0vy-CmsuMg7l4V4s2l-Bz2iqXrgyzBl0q3G4wmlJ1rJr-z_Ks40pRhzzhQxjttDp-r_B47l_Bxg_4E-00oC4yhd458iE-_ojD72owEnmjzV8xr_C0hp-K9p16Bt17LrhrZ_imJ9lnWq0mtB-902NgivwD2qx8Bv5s4Gug1kUsl4_Cm5gJt8_kC4tmU7ougB6oyjE6tvJ2573Biv4kBq54Km1mwDsr2DquvK-i87C98i6E6zv4B9hlI-z7-Fl2vrDto9kB87szEh_n6G9nn2I_zzpL5-zwFz6_f8lzF6tquBonzciv2oBp-pMw69f9wzG8v_lCxzw3Bk0yuCu07iBq5zsDok5_CwlkqC_2-rD2_knDv20iCotwQ6tuvC2y0-DpuvkBv74L6m1jB6nnwIh-qxBxzzN1_1nCow6Egzs_I3n5kF_sooOy9_-Li2r7Bum5xGvz8Iu2rnFjzlre3nBlq84VpisynByvttR3lvhK04-0F7puxJlx3pNxgv7O4qxjI0np9vBwp1nHso4oMhqvkFmo9zEp1_-B010e2unwBwk2kb86jsbl9kyJnnuhJ_z1mL6wz6Sr_w7QnvlzYkzk8D422jJgttnFg_w6N1i_lM9x39J34zw4Bi1hlD4g54D01oxFw_pzNm53kTpvzmQ3vpqYhxk_F7o01Zl91rHz0mwWk989Mn6hsBsrl2E-x7yCxl5iFzk3kC1-olF1u4nG69lzHz4j8Otpm3Mhg7oI_0hhE2vwjH0m0gJypx6Q9to5D93pkK87z0Q971lR6ms-CowvxG2_08Hn9skIu3lzGxm85Dk3iiJ530pFvwyhwBixkxYs62_OsshgX3q6hL6lpzUi65sNytrs1B3whjjBkhj4bsg85Q-ixtDqvq3Cnl2lZ2yy7SqthmrBw1n2H35nrLq9xowB6x-mS8ix6X-3wlDvsz0B7ilzDq61uG56xkEo19mBx7u0Etk2rC_-1pEw9lxP86vhX-2m8O2v9kG1h2rNpuq2Ts_xiUioqxR6-xuPvi6nM0wzxSz342Z05yiVno17CpzohHpq95Cq3jvCjp2hBq5w_Uq51jvC2iztmBx4w-W7lwnQsny2Nl8n8Go-w7D9un1F5g5_Hi90wMmj6qLqs79OyqhjIj2qjKlsusLn5h2L4g1pHy11wDxp9vGws_uHqj8tGli1yC-v8sDp3plC_gq8oG65urpBngm-cs9z9Fv_t4L4mrklBq3ns0Bxv5wlB0iqpmB9g3_ag_unV89_2G1s6vCks1_cl8jyDp667MioypN5rx0a43mqIu63lB1gl-Dv67clmliCqownErss7Cx9grB8nyzZ48lmCi8g9Cy_nyJ5-s9Ci2j6Frrxsd_92kKxm3wOi-5mXwg0jEylj4Fhr6lDsr00Fmju3D6x5sD6um8B9k5kC_si3Cshp_Bls0RjkxnBnthH35huCq4quG9_orDnmpR8ih2BwmhdlqsqB8t23Ep1xJ1goqBsm4tC66sGrnopCqj5tEly_2Bp1qFxnzsH-hxzQ0wzjCs97f2qz1CjginGwjihDt94Or_tI3kvgK9xqbjl36B-qs-Gpyx2J10-oK7u4Vk1opBjms4Bsptej9qbgquIng7kBtkjwBtovvB6wo_Dhqn0DxuoiG2rgmCr4owB8g8Txl6Tj-7Fyq-gD5rtqB9goqCs4_lB8k1Ewn7TkxzyBwm49B714tDxphoDg_12DupjpBpgs4BlnpiDw5yH06_0C7h0yBtkz6C-_o2Cy6_oI85xtD43jgB089rDzu9_CqqjV2qw7BtlxwCj-5sNwlsnB3w7fuknkEui_rD50kW9rvqClr_9Duni2BghzqCu93tFw2_3IpgpxHmz4wBok1wCkk7rDyu9lM1ixhDgnjjEz9opBprm3B143vCgk58B-tslDpn8hEkm2uB_rxkBynv1Cz1zoC-5nwCsmM2w5xCk_mfumlTmw8evwnSspwlB1o83V776qExsxqEl8yxE463zCourmEv7ojB45rnBggg8D90puBr67R7nm3Hzh8nQ2hr9Lo6p2Bl1khI0_7qM9g2lIum02Vhz3gP59ioGpu0Gjmj-Es2r0CrkknD6sorOgywlH151wDy96vC5y_4BlvkmE2y1oUz9qiJ1kipGx1q9DynhqE9lxjBt7_Wn6hYkwokB3ms_Bx_rhC7v1f5r56Evk76cvg9pB-pka96zd6yu1Cj28uCry5oBh_o3F1ny9C1nqZn2m7Cjngfoir-C3hlYv6-Gl367Np4yoCww0jFotk8Cs6ooGil4nO8t6Qq7ijCn39Kv8pNgiygK9l38Mq6rvdg1_7pBvs5pwBmu45f93skdj1xrF2j06DtpowCn2sqCuv-nN49noL-2_7Ep71mFzxxkGr49sJ4k4nKx_7uHg-l8Gysu0Fw5h7CyoohEuo46C7gsgD5wv2CisimF5vloI-1iyL520H-sl-B9xijB_zzmekwyjT5t38KlupqIut85Eoos2Bix9-Bp87jBzzqHszkiBloorOwxlGiupaz2soBtjyXr2_Ozx-XqswPj-kbvugM_6jzEny2mDigrlB88pPm09yF3-5Nq8vhBt9hsSnpuqBq9ikHll3hB-s9fn8m6Dxy6qDrku3Ckv61Fm2qrDjnu4Bn-9jDxnlsD38-iC6virDu-54CgwirB484_Bw07U593_D8koxB6l20Kujt8C18-iCi-jjIrnvlDrqwxFrix8U59nrD83l7Do0oiFm-vzEvrjgBlj2zCh5-Xt1hrBz4qkB6qx8Dzs9-G586qCihnxK7747Elgi0Cn8sTkrqhCt94iBxolD_82Yz-9_Bli1VzrshBu_xkFy0lzGowwhGik08O5njgBlw_wN_qs9Jp91L_z4jiDq-goKh73xL0x5qG-gslI5qoS6r44BjuzV00owC6mmF06wDpp4jU1twxHjlx3Eoo7tB6q45F0m1iL-655E09qsC762L_5kL41x5B-t3c3o7jBpg2nC5ysFv9jM_roG86kiBt7zmJ5wymB79hxDnqysOh8g2Cmmh3D-pl9Cg_03D96uwDjs3jC9q-2Bn05rEozohB-j2_Fg581hBy7z0S5rv1H85_rBon3r9C9mmbqyqySwmuyO35qqCnjmsFor3zG0xz9Lw_5mHqjg6B6h-uK11ktY-rqqE52ylCvzjqF75u_Dmr_9Dz-kvDu47T64mxB214yE7z5jEr1-uGuo5xOvh2vSwwkgW5-voZvjpvbu0w-evz5wlBxx57X1g28Goj8vHxus8D3uvzGzrwlKog4Ls_sdz__oBx229Etng5E-gppDo_unBysm2H6j5lK_30gD1ptkEno6yB4x0uHk6hc2h1uBzo0GrwyzB4y08OgxvhFt3rzBgg0kCls6xFms48JyljpD5ygvBqj7iCyrtRoy9Rk14iJ09q4Mi8lqC6mp3JwrvmEg-6iC5u2kC4-48Bj51hBr-hqGzzmrBosskCq4xKgpt7gB3o3gBzz_2B_m8Vrk6gBt8tQ1y4Rjr1W4viNg-uezh4Sou87BkmmtB4t-Mio2nBvwt4Bhkw4B6u1a14ngBpyueyir1Hq6ovJqg_9LlpgxGjx9pCx17nFqigN7i7Ym9nfqzuyD1thgE3r7mDmx20Jtls_Fys_Svyy6Htx9yFlothEt73gM3mzUz6ulEyu0pEnj7pB4g3zD8iynBj950Bln31BxxlGhwmhBunssC0t0lDzyvgKgxxTl-yoB_5ipChztwCvlxgB_tioJhj85Cr7ykHq98tE0wm2Hugw_Jg_mmDjjrR8os3D876pG0g8oBzotoB2mprB0szjEt6__Cm92Vlp-pCzyn9Bzt-_Gtgv1Ckl2oBtx-qBihskEwj5tBojxyBoqid9s5nBnlp0C-6wkBznnQoywiEj4wkDpx8nCvv_kF0ukqB03lrB069tEqqugB0t0nI9sw4Kypz7G0ivqDtluxBoh3-Dkir-BtxxgD7tj_Gx-q2Es_h-Eph6hHgx1gD8h-wD4g78MvhwkI_oy4Ftol5EqviyC862zC-ru_Fpu6_Bv6ioKm71choy8Bk6ptEn8_xBuknuFjmwM5g--D5wm-F2vtkErzxhHkz-hI2tx0Bz7hkDk0mhBpvptF5h10B2uzX7jqpBgl-Yn4pkD5is0B2q7a-7wiB8yoyDov6vBjkxsGp9i5Bi7orB9p87B0s-Lh8m5G0_-oDntxhC5s0c5jnxEj-qfo-vdj1hzEq9j1Q_12vHp32Uo2y2Fk6y6Dy92xC9ypepqhyB_pjY7sunB3lubghxXiikpBy9ruCqxywT0ou2Lij_uEyjptFnvk_Cz-vuC4o8hD68itFs9rM28qwBgun7uB3p96Op6kwCmkmnEts9J-_mIu5vM92qU3mx3DrxpjO5jt4B2wyVvm-wCsq2xDk0vwCtz4Kq7-Ogm0d6_g6C4pxV0vijC13v6Bnj4VitozBuh6hCm-3Cy7g8C6q-tBxt5pBln9_B3nwd3qjP4p6wBpmjxE9rilCw0uT8imRko0Sv_4oB52moE1v2sBq8q8Eh-0mCy964Cig4Tp84zBqk2_E4r50BkhvtCx2v0Bv0n4EphpgDr0nRhk9Dt5_gHzl73ErlkzBqp5P3jzf-3u3CzjomY1sw0HnwimG3wipGs5hlQjv3jBp-7qC9gzmBgl32lB_wpzB4t7kJp62Zthp-Exw2oGtg2S8t8uC-umxBjxzzF19hyHqg8wFog0-Izo_0Gj2m0Dn0n4E5l7qC_ui-E9-mtap7lnYhv6pa_zx9Ypo9pXu8q3Ou-iuF4y5-G47j1Cu5_C810Bhy13B7m-lGymv_F_mzhD23z0Ejss6HiugsG-1_4Dsgq0Fm60kBqxkrFtl7kD-szkM_pt-Mo6jyTgqi2G0ykyBm2nsNw15kK5li0EqtorJ13skI-zzwH066rJoyoN_0utH5-hvFmrnnEms_-wByg8hbpoyuoBq2-4Tik30Sp4h5I9x02Fq-tO4xwQ-wusBt40W30_oCu2n0G72sckkv2Czkr7Frm69EqxiZw7l0CyujnBnl_T16wyEn23uEk85-c335hG_jhL0662C4swrC00soB34h3F8gn2E0oj8Vqt8yH6plpCgkuXlyk6K71gwBw75-B6-kjBvs0Nhy40Bt53L9sx8Blp9gEi7n3oB-v4jYmo-nQo_j6b79m5Rw4g4Oqu2lT8vz15B4y6sWmqvnU_gixGur2jB0wx9M1jj1Hit7wJjkgtE8-xqFxy1qD3yhCln8mZ8zwgR0_pwJ-0rhGiqiRw0rrkD",
          Target: "https://www.google.ca",
          Title: "Title"
        }
    ];
    
    this._searchResults = {
      RelevantResults: results,
      QueryKeywords: "Test",
      RefinementResults:[]
    };   

  }

  public render(): void {

    if (Environment.type === EnvironmentType.Local) {
      this.setMockResults();
    }
    
    let center: number[] = [ this.tryParseFloat(this.properties.centerLatitude), this.tryParseFloat(this.properties.centerLongitude) ];
    this._map = React.createElement(
      BingMaps, { 
        componentId: this.componentId, 
        polygonResults: this._searchResults,
        templateService: this._templateService,
        bingMapsAPIKey: this.properties.bingMapsAPIKey,
        mapTypeId: this.properties.mapTypeId,
        zoom: this.properties.zoom,
        navigationBarMode: this.properties.navigationBarMode,
        supportedMapTypes: this.properties.supportedMapTypes,
        columns: this.properties.columns,
        titleColumnName: this.properties.showLabels == false ? "" : this.properties.titleColumnName,
        polygonColumnName: this.properties.polygonColumnName,
        targetColumnName: this.properties.targetColumnName,
        fillColor: this.toColorString(this.properties.fillColor),
        strokeColor: this.toColorString(this.properties.strokeColor),
        strokeThickness: this.properties.strokeThickness,
        center: center,
        height: this.properties.height,
        startOptions: {
          disableStreetside:this.properties.disableStreetside,
          showDashboard:this.properties.showDashboard,
          showLocateMeButton: this.properties.showLocateMeButton,
          showMapTypeSelector:this.properties.showMapTypeSelector,
          showScalebar:this.properties.showScalebar,
          showZoomButtons:this.properties.showZoomButtons
        },
        mapOptions: {
          disableZooming:this.properties.disableZooming,
          disableScrollWheelZoom:this.properties.disableScrollWheelZoom,
          allowInfoboxOverflow:this.properties.allowInfoboxOverflow,
          disableBirdseye:this.properties.disableBirdseye,
          disablePanning:this.properties.disablePanning,
          maxZoom: this.properties.maxZoom,
          minZoom: this.properties.minZoom
        }
      }
    );

    this._component = ReactDom.render(this._map, this.domElement) as BingMaps;
    
  }

  private toColorString(val:any) : string {
    if(!val) return val;
    if(typeof val == "string") return val;
    if(typeof val == "object") {
      return 'rgba(' + val.r.toString() + ',' + val.g.toString() + ',' + val.b.toString() + ',' + val.a.toString() + ')';
    }
  }

  private tryParseFloat(val: string) : number {
    try {
      return parseFloat(val);
    } catch(ex){
      return 0;
    }
  }

  public onChangeHappened(e: ISearchEvent) {
    
    console.log("Recieved search results");
    console.log(e.results);

    if(this._map) this._map.props.polygonResults = e.results;
    if(this._component) this._component.forceUpdate();

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected async onPropertyPaneConfigurationStart() {
    await this.loadPropertyPaneResources();
  }

  protected async loadPropertyPaneResources(): Promise<void> {

    // tslint:disable-next-line:no-shadowed-variable
    const { PropertyFieldCodeEditor, PropertyFieldCodeEditorLanguages } = await import(
        /* webpackChunkName: 'search-property-pane' */
        '@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor'
    );

    this._propertyFieldCodeEditor = PropertyFieldCodeEditor;
    this._propertyFieldCodeEditorLanguages = PropertyFieldCodeEditorLanguages;
      
  }

  /**
     * Checks if a field if empty or not
     * @param value the value to check
     */
    private _validateEmptyField(value: string): string {

      if (!value) {
          return "This is a required field.";
      }

      return '';
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        this.getPolygonOptionsPage(),
        this.getColorsPage(),
        this.getBingMapsOptionsPage()
      ]
    };
  }

  protected getColorsPage():IPropertyPanePage {
    return {
      header: {
        description:strings.ColorsPageDescription
      },
      groups: [
        {
          groupName: strings.colorsStylesGroupLabel,
          groupFields: [
            PropertyPaneToggle('showLabels', {
                label: strings.showLabelsLabel, 
                checked: this.properties.showLabels
            }), 
            PropertyFieldNumber('strokeThickness', {
              key: "strokeThickness",
              label: strings.strokeThicknessLabel,
              value: this.properties.strokeThickness,
              onGetErrorMessage: this._validateEmptyField.bind(this),
              minValue: 1,
              maxValue: 10
            }),
            PropertyFieldColorPicker('fillColor', {
              label: strings.fillColorLabel,
              properties: this.properties,
              selectedColor: this.properties.fillColor,
              onPropertyChange: (e) => {},
              style: PropertyFieldColorPickerStyle.Full,
              valueAsObject: true,
              iconName: 'BucketColor',
              key: 'fillColorID'
            }),
            PropertyFieldColorPicker('strokeColor', {
              label: strings.strokeColorLabel,
              properties: this.properties,
              selectedColor: this.properties.strokeColor,
              onPropertyChange: (e)=> {},
              style: PropertyFieldColorPickerStyle.Full,
              valueAsObject: true,
              iconName: 'Line',
              key: 'strokeColorID'
            })
          ]
        }
      ]
    };
  }

  private getBingMapsOptionsPage():IPropertyPanePage {
    return {
      header: {
        description:strings.BingMapsPageDescription
      },
      groups: [
        {
          groupName: strings.bingMapsGroupNameExtendedLabel,
          groupFields: [
            PropertyPaneDropdown('mapTypeId', {
              label: strings.mapTypeIdLabel,
              selectedKey: this.properties.mapTypeId,
              options: this._mapTypeIds
            }),
            PropertyFieldMultiSelect('supportedMapTypes', {
              key: 'supportedMapTypes',
              label: strings.supportedMapTypesLabel,
              options: this._mapTypeIds,
              selectedKeys: this.properties.supportedMapTypes
            }),
            PropertyPaneToggle('disableZooming', {
              label: strings.disableZoomingLabel, 
              checked: this.properties.disableZooming
            }), 
            PropertyPaneToggle('disableScrollWheelZoom', {
              label: strings.disableScrollWheelZoomLabel, 
              checked: this.properties.disableScrollWheelZoom
            }), 
            PropertyPaneToggle('allowInfoboxOverflow', {
              label: strings.allowInfoboxOverflowLabel, 
              checked: this.properties.allowInfoboxOverflow
            }), 
            PropertyPaneToggle('disableBirdseye', {
              label: strings.disableBirdseyeLabel, 
              checked: this.properties.disableBirdseye
            }), 
            PropertyPaneToggle('disablePanning', {
              label: strings.disablePanningLabel, 
              checked: this.properties.disablePanning
            }), 
            PropertyPaneToggle('disableStreetside', {
              label: strings.disableStreetsideLabel, 
              checked: this.properties.disableStreetside
            }), 
            PropertyPaneToggle('showDashboard', {
              label: strings.showDashboardLabel, 
              checked: this.properties.showDashboard
            }), 
            PropertyPaneToggle('showLocateMeButton', {
              label: strings.showLocateMeButtonLabel, 
              checked: this.properties.showLocateMeButton
            }), 
            PropertyPaneToggle('showMapTypeSelector', {
              label: strings.showMapTypeSelectorLabel, 
              checked: this.properties.showMapTypeSelector
            }), 
            PropertyPaneToggle('showScalebar', {
              label: strings.showScalebarLabel, 
              checked: this.properties.showScalebar
            }), 
            PropertyPaneToggle('showZoomButtons', {
              label: strings.showZoomButtonsLabel, 
              checked: this.properties.showZoomButtons
            }),
            PropertyFieldNumber("minZoom", {
              key: "minZoom",
              label: strings.minZoomLabel,
              value: this.properties.minZoom,
              minValue: 1,
              maxValue: 19
            }),
            PropertyFieldNumber("maxZoom", {
              key: "maxZoom",
              label: strings.maxZoomLabel,
              value: this.properties.maxZoom,
              minValue: 1,
              maxValue: 19
            })
          ]
        },
      ]
    };
  }

  private getPolygonOptionsPage() : IPropertyPanePage {
    return {
      header: {
        description: strings.PropertyPaneDescription
      },
      groups: [
        {
          groupName: strings.bingMapsGroupNameLabel,
          groupFields: [
            PropertyPaneTextField('bingMapsAPIKey', {
              label: strings.bingMapsAPIKeyLabel,
              value: this.properties.bingMapsAPIKey,
              onGetErrorMessage: this._validateEmptyField.bind(this)
            }),
            PropertyPaneTextField('height', {
              label: strings.mapHeightLabel,
              value: this.properties.height
            }),
            PropertyFieldNumber('zoom', {
              key: "zoom",
              label: strings.zoomLabel,
              value: this.properties.zoom,
              onGetErrorMessage: this._validateEmptyField.bind(this),
              minValue: 1,
              maxValue: 19
            }),
            PropertyPaneTextField('centerLatitude', {
              label: strings.centerLatitudeLabel,
              value: this.properties.centerLatitude.toString(),
              onGetErrorMessage: this._validateEmptyField.bind(this),
            }),
            PropertyPaneTextField('centerLongitude', {
              label: strings.centerLongitudeLabel,
              value: this.properties.centerLongitude.toString(),
              onGetErrorMessage: this._validateEmptyField.bind(this)
            }),
            
          ]
        },
        {
          groupName: strings.columnConfigurationGroupLabel,
          groupFields: [
            PropertyPaneTextField('titleColumnName', {
              label: strings.titleColumnNameLabel,
              value: this.properties.titleColumnName,
              onGetErrorMessage: this._validateEmptyField.bind(this)
            }),
            PropertyPaneTextField('polygonColumnName', {
              label: strings.polygonColumnNameLabel,
              value: this.properties.polygonColumnName
            }),
            PropertyPaneTextField('targetColumnName', {
              label: strings.targetColumnNameLabel,
              value: this.properties.targetColumnName
            }),
            PropertyFieldCollectionData('columns',{
              key: 'columns',
              label: strings.columnsLabel,
              panelHeader: strings.columnsPanelHeader,
              manageBtnLabel: strings.columnsButtonLabel,
              value: this.properties.columns,
              enableSorting: true,
              disableItemCreation: true,
              disableItemDeletion: true,
              fields: [{
                  id:"name",
                  title:strings.columnsNameColumnTitle,
                  type: CustomCollectionFieldType.string,
                  required: true
                }]
            })
          ]
        }        
      ]
    };

  }

}
