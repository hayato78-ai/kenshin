/**
 * UIMapping.js - UI表示層
 *
 * 設計原則:
 * - エンティティオブジェクト ⇔ UI表示用オブジェクトの変換を担当
 * - UIの表示ラベルや形式はここで定義
 * - 物理的な列位置は参照しない（DAOに委譲）
 *
 * @version 1.0.0
 * @date 2025-12-22
 */

// ============================================
// UIMapping - UI表示マッピング
// ============================================

const UIMapping = {

  // ============================================
  // 受診者データ変換
  // ============================================

  /**
   * エンティティ → UI表示用オブジェクトに変換
   * @param {Object} entity - DAOから取得したエンティティ
   * @returns {Object} UI表示用オブジェクト
   */
  patientToUI(entity) {
    if (!entity) return null;

    return {
      // 識別情報
      patientId: entity.patientId || '',
      karteNo: entity.karteNo || '',

      // 表示用ラベル付きフィールド
      display: {
        patientId: entity.patientId || '-',
        karteNo: entity.karteNo || '-',
        status: entity.status || '-',
        visitDate: this.formatDateForDisplay(entity.visitDate),
        name: entity.name || '-',
        kana: entity.kana || '-',
        gender: entity.gender || '-',
        birthdate: this.formatDateForDisplay(entity.birthdate),
        age: entity.age !== '' ? String(entity.age) + '歳' : '-',
        course: entity.course || '-',
        company: entity.company || '-',
        department: entity.department || '-',
        overallJudgment: entity.overallJudgment || '-'
      },

      // フォーム用（編集可能なフィールド）
      form: {
        patientId: entity.patientId || '',
        karteNo: entity.karteNo || '',
        status: entity.status || '入力中',
        visitDate: this.formatDateForForm(entity.visitDate),
        name: entity.name || '',
        kana: entity.kana || '',
        gender: entity.gender || '',
        birthdate: this.formatDateForForm(entity.birthdate),
        age: entity.age || '',
        course: entity.course || '',
        company: entity.company || '',
        department: entity.department || '',
        overallJudgment: entity.overallJudgment || ''
      },

      // 内部用メタ情報
      _rowIndex: entity._rowIndex
    };
  },

  /**
   * UI入力 → エンティティに変換
   * @param {Object} uiData - UIからの入力データ
   * @returns {Object} エンティティオブジェクト
   */
  uiToPatient(uiData) {
    if (!uiData) return null;

    return {
      patientId: uiData.patientId || '',
      karteNo: uiData.karteNo || '',
      status: uiData.status || '入力中',
      visitDate: this.parseFormDate(uiData.visitDate || uiData.examDate),
      name: uiData.name || '',
      kana: uiData.kana || uiData.nameKana || '',
      gender: uiData.gender || '',
      birthdate: this.parseFormDate(uiData.birthdate || uiData.birthDate),
      age: this.calculateAge(uiData.birthdate || uiData.birthDate),
      course: uiData.course || '',
      company: uiData.company || '',
      department: uiData.department || '',
      overallJudgment: uiData.overallJudgment || '',
      csvImportDate: uiData.csvImportDate || '',
      exportDate: uiData.exportDate || '',
      bmlPatientId: uiData.bmlPatientId || '',
      _rowIndex: uiData._rowIndex || null
    };
  },

  /**
   * 一覧表示用の配列に変換
   * @param {Array<Object>} entities - エンティティ配列
   * @returns {Array<Object>} UI表示用配列
   */
  patientsToUIList(entities) {
    if (!entities || !Array.isArray(entities)) return [];
    return entities.map(entity => this.patientToUI(entity));
  },

  // ============================================
  // 検索結果変換
  // ============================================

  /**
   * 検索結果をUI用に変換
   * @param {Array<Object>} entities - 検索結果エンティティ配列
   * @returns {Object} UI表示用検索結果
   */
  searchResultsToUI(entities) {
    return {
      success: true,
      count: entities ? entities.length : 0,
      results: this.patientsToUIList(entities),
      timestamp: new Date().toISOString()
    };
  },

  // ============================================
  // 日付フォーマット
  // ============================================

  /**
   * 表示用日付フォーマット（YYYY/MM/DD）
   * @param {*} value - 日付値
   * @returns {string} フォーマットされた日付
   */
  formatDateForDisplay(value) {
    if (!value) return '-';
    try {
      const date = value instanceof Date ? value : new Date(value);
      if (isNaN(date.getTime())) return '-';
      const y = date.getFullYear();
      const m = String(date.getMonth() + 1).padStart(2, '0');
      const d = String(date.getDate()).padStart(2, '0');
      return `${y}/${m}/${d}`;
    } catch (e) {
      return '-';
    }
  },

  /**
   * フォーム用日付フォーマット（YYYY-MM-DD）
   * @param {*} value - 日付値
   * @returns {string} フォーマットされた日付
   */
  formatDateForForm(value) {
    if (!value) return '';
    try {
      const date = value instanceof Date ? value : new Date(value);
      if (isNaN(date.getTime())) return '';
      const y = date.getFullYear();
      const m = String(date.getMonth() + 1).padStart(2, '0');
      const d = String(date.getDate()).padStart(2, '0');
      return `${y}-${m}-${d}`;
    } catch (e) {
      return '';
    }
  },

  /**
   * フォーム入力の日付をDateに変換
   * @param {string} value - YYYY-MM-DD形式の日付
   * @returns {Date|null} Dateオブジェクトまたはnull
   */
  parseFormDate(value) {
    if (!value) return null;
    try {
      const date = new Date(value);
      if (isNaN(date.getTime())) return null;
      return date;
    } catch (e) {
      return null;
    }
  },

  // ============================================
  // 年齢計算
  // ============================================

  /**
   * 生年月日から年齢を計算
   * @param {*} birthdate - 生年月日
   * @returns {number|string} 年齢または空文字
   */
  calculateAge(birthdate) {
    if (!birthdate) return '';
    try {
      const birth = birthdate instanceof Date ? birthdate : new Date(birthdate);
      if (isNaN(birth.getTime())) return '';

      const today = new Date();
      let age = today.getFullYear() - birth.getFullYear();
      const monthDiff = today.getMonth() - birth.getMonth();

      if (monthDiff < 0 || (monthDiff === 0 && today.getDate() < birth.getDate())) {
        age--;
      }

      return age >= 0 ? age : '';
    } catch (e) {
      return '';
    }
  },

  // ============================================
  // ステータス変換
  // ============================================

  /**
   * ステータスの表示用ラベル取得
   * @param {string} status - ステータス値
   * @returns {Object} 表示情報
   */
  getStatusDisplay(status) {
    const statusMap = {
      '入力中': { label: '入力中', color: 'warning', icon: '📝' },
      '入力完了': { label: '入力完了', color: 'success', icon: '✅' },
      '確認済': { label: '確認済', color: 'info', icon: '👁️' },
      '出力済': { label: '出力済', color: 'secondary', icon: '📤' }
    };
    return statusMap[status] || { label: status || '-', color: 'default', icon: '❓' };
  },

  // ============================================
  // バリデーション
  // ============================================

  /**
   * UI入力のバリデーション
   * @param {Object} uiData - UIからの入力データ
   * @returns {Object} バリデーション結果
   */
  validatePatientInput(uiData) {
    const errors = [];

    // 必須チェック
    if (!uiData.name || !uiData.name.trim()) {
      errors.push({ field: 'name', message: '氏名は必須です' });
    }

    if (!uiData.birthdate && !uiData.birthDate) {
      errors.push({ field: 'birthdate', message: '生年月日は必須です' });
    }

    // 形式チェック
    if (uiData.karteNo && !/^\d+$/.test(uiData.karteNo)) {
      errors.push({ field: 'karteNo', message: 'カルテNoは数字のみです' });
    }

    // 日付チェック
    const birthValue = uiData.birthdate || uiData.birthDate;
    if (birthValue) {
      const date = new Date(birthValue);
      if (isNaN(date.getTime())) {
        errors.push({ field: 'birthdate', message: '生年月日の形式が不正です' });
      } else if (date > new Date()) {
        errors.push({ field: 'birthdate', message: '生年月日は未来日にできません' });
      }
    }

    return {
      valid: errors.length === 0,
      errors: errors
    };
  },

  // ============================================
  // UI定義（ラベル・プレースホルダー）
  // ============================================

  /**
   * フィールド定義を取得
   * @returns {Object} フィールド定義
   */
  getFieldDefinitions() {
    return {
      patientId: { label: '受診者ID', placeholder: '自動生成', editable: false },
      karteNo: { label: 'カルテNo', placeholder: '例: 999991', editable: true },
      status: { label: 'ステータス', placeholder: '', editable: true },
      visitDate: { label: '受診日', placeholder: 'YYYY-MM-DD', editable: true },
      name: { label: '氏名', placeholder: '例: 山田太郎', editable: true, required: true },
      kana: { label: 'カナ', placeholder: '例: ヤマダタロウ', editable: true },
      gender: { label: '性別', placeholder: '男/女', editable: true },
      birthdate: { label: '生年月日', placeholder: 'YYYY-MM-DD', editable: true, required: true },
      age: { label: '年齢', placeholder: '自動計算', editable: false },
      course: { label: '受診コース', placeholder: '', editable: true },
      company: { label: '事業所名', placeholder: '', editable: true },
      department: { label: '所属', placeholder: '', editable: true },
      overallJudgment: { label: '総合判定', placeholder: '', editable: true }
    };
  },

  /**
   * ステータス選択肢を取得
   * @returns {Array<Object>} 選択肢配列
   */
  getStatusOptions() {
    return [
      { value: '入力中', label: '入力中' },
      { value: '入力完了', label: '入力完了' },
      { value: '確認済', label: '確認済' },
      { value: '出力済', label: '出力済' }
    ];
  },

  /**
   * 性別選択肢を取得
   * @returns {Array<Object>} 選択肢配列
   */
  getGenderOptions() {
    return [
      { value: '男', label: '男' },
      { value: '女', label: '女' }
    ];
  }
};
