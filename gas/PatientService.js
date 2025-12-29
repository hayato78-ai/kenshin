/**
 * PatientService.js - ビジネスロジック層
 *
 * 設計原則:
 * - ビジネスルール（重複チェック、ステータス遷移等）を担当
 * - DAOを使ってデータアクセス
 * - UIMappingを使ってUIとの変換
 * - portalApi.jsから呼び出される
 *
 * @version 1.0.0
 * @date 2025-12-22
 */

// ============================================
// PatientService - ビジネスロジック
// ============================================

const PatientService = {

  // ============================================
  // 受診者登録
  // ============================================

  /**
   * 受診者を新規登録
   * @param {Object} uiData - UIからの入力データ
   * @returns {Object} 結果 {success, patientId, message, errors}
   */
  registerPatient(uiData) {
    try {
      // 1. バリデーション
      const validation = UIMapping.validatePatientInput(uiData);
      if (!validation.valid) {
        return {
          success: false,
          errors: validation.errors,
          message: validation.errors.map(e => e.message).join(', ')
        };
      }

      // 2. UI → エンティティ変換
      const entity = UIMapping.uiToPatient(uiData);

      // 3. 重複チェック（氏名＋生年月日）
      if (entity.name && entity.birthdate) {
        const existing = PatientDAO.getByNameAndBirthdate(entity.name, entity.birthdate);
        if (existing) {
          return {
            success: false,
            message: `同一の氏名・生年月日の受診者が既に存在します（ID: ${existing.patientId}）`,
            existingPatient: UIMapping.patientToUI(existing)
          };
        }
      }

      // 4. カルテNo重複チェック（指定されている場合）
      if (entity.karteNo) {
        const existingByKarte = PatientDAO.getByKarteNo(entity.karteNo);
        if (existingByKarte) {
          return {
            success: false,
            message: `カルテNo ${entity.karteNo} は既に使用されています（ID: ${existingByKarte.patientId}）`
          };
        }
      }

      // 5. デフォルト値設定
      if (!entity.status) {
        entity.status = '入力中';
      }

      // 6. 保存
      const result = PatientDAO.save(entity);

      if (result.success) {
        logInfo(`PatientService.registerPatient: 登録成功 ${result.patientId}`);
        return {
          success: true,
          patientId: result.patientId,
          message: `受診者を登録しました（ID: ${result.patientId}）`
        };
      } else {
        return {
          success: false,
          message: result.error || '登録に失敗しました'
        };
      }

    } catch (e) {
      logError('PatientService.registerPatient', e);
      return {
        success: false,
        message: 'システムエラー: ' + e.message
      };
    }
  },

  // ============================================
  // 受診者更新
  // ============================================

  /**
   * 受診者情報を更新
   * @param {Object} uiData - UIからの入力データ（patientId必須）
   * @returns {Object} 結果 {success, message, errors}
   */
  updatePatient(uiData) {
    try {
      if (!uiData.patientId) {
        return {
          success: false,
          message: '受診者IDが指定されていません'
        };
      }

      // 1. 存在確認
      const existing = PatientDAO.getById(uiData.patientId);
      if (!existing) {
        return {
          success: false,
          message: '指定された受診者が見つかりません'
        };
      }

      // 2. バリデーション
      const validation = UIMapping.validatePatientInput(uiData);
      if (!validation.valid) {
        return {
          success: false,
          errors: validation.errors,
          message: validation.errors.map(e => e.message).join(', ')
        };
      }

      // 3. UI → エンティティ変換
      const entity = UIMapping.uiToPatient(uiData);
      entity._rowIndex = existing._rowIndex;

      // 4. カルテNo重複チェック（変更時）
      if (entity.karteNo && entity.karteNo !== existing.karteNo) {
        const existingByKarte = PatientDAO.getByKarteNo(entity.karteNo);
        if (existingByKarte && existingByKarte.patientId !== entity.patientId) {
          return {
            success: false,
            message: `カルテNo ${entity.karteNo} は既に使用されています`
          };
        }
      }

      // 5. 更新
      const result = PatientDAO.update(entity);

      if (result.success) {
        logInfo(`PatientService.updatePatient: 更新成功 ${entity.patientId}`);
        return {
          success: true,
          message: '受診者情報を更新しました'
        };
      } else {
        return {
          success: false,
          message: result.error || '更新に失敗しました'
        };
      }

    } catch (e) {
      logError('PatientService.updatePatient', e);
      return {
        success: false,
        message: 'システムエラー: ' + e.message
      };
    }
  },

  // ============================================
  // 受診者検索
  // ============================================

  /**
   * 受診者を検索
   * @param {Object} criteria - 検索条件
   * @returns {Object} 検索結果 {success, count, results}
   */
  searchPatients(criteria) {
    try {
      const entities = PatientDAO.search(criteria);
      return UIMapping.searchResultsToUI(entities);
    } catch (e) {
      logError('PatientService.searchPatients', e);
      return {
        success: false,
        count: 0,
        results: [],
        message: 'システムエラー: ' + e.message
      };
    }
  },

  /**
   * 受診者をIDで取得
   * @param {string} patientId - 受診者ID
   * @returns {Object|null} UI表示用オブジェクト
   */
  getPatientById(patientId) {
    try {
      const entity = PatientDAO.getById(patientId);
      return entity ? UIMapping.patientToUI(entity) : null;
    } catch (e) {
      logError('PatientService.getPatientById', e);
      return null;
    }
  },

  /**
   * 受診者をカルテNoで取得
   * @param {string} karteNo - カルテNo
   * @returns {Object|null} UI表示用オブジェクト
   */
  getPatientByKarteNo(karteNo) {
    try {
      const entity = PatientDAO.getByKarteNo(karteNo);
      return entity ? UIMapping.patientToUI(entity) : null;
    } catch (e) {
      logError('PatientService.getPatientByKarteNo', e);
      return null;
    }
  },

  /**
   * 全受診者を取得（ページング対応）
   * @param {Object} options - オプション {page, limit, sortBy, sortOrder}
   * @returns {Object} 結果 {success, count, results, pagination}
   */
  getAllPatients(options = {}) {
    try {
      const page = options.page || 1;
      const limit = options.limit || 50;

      const allEntities = PatientDAO.getAll();
      const total = allEntities.length;

      // ソート
      if (options.sortBy) {
        allEntities.sort((a, b) => {
          const valA = a[options.sortBy] || '';
          const valB = b[options.sortBy] || '';
          const comp = String(valA).localeCompare(String(valB), 'ja');
          return options.sortOrder === 'desc' ? -comp : comp;
        });
      }

      // ページング
      const start = (page - 1) * limit;
      const end = start + limit;
      const pagedEntities = allEntities.slice(start, end);

      return {
        success: true,
        count: pagedEntities.length,
        total: total,
        results: UIMapping.patientsToUIList(pagedEntities),
        pagination: {
          page: page,
          limit: limit,
          totalPages: Math.ceil(total / limit),
          hasNext: end < total,
          hasPrev: page > 1
        }
      };

    } catch (e) {
      logError('PatientService.getAllPatients', e);
      return {
        success: false,
        count: 0,
        total: 0,
        results: [],
        message: 'システムエラー: ' + e.message
      };
    }
  },

  // ============================================
  // ステータス管理
  // ============================================

  /**
   * ステータスを更新
   * @param {string} patientId - 受診者ID
   * @param {string} newStatus - 新しいステータス
   * @returns {Object} 結果
   */
  updateStatus(patientId, newStatus) {
    try {
      // ステータス遷移チェック
      const validStatuses = ['入力中', '入力完了', '確認済', '出力済'];
      if (!validStatuses.includes(newStatus)) {
        return {
          success: false,
          message: '無効なステータスです'
        };
      }

      const result = PatientDAO.updateFields(patientId, { status: newStatus });

      if (result.success) {
        logInfo(`PatientService.updateStatus: ${patientId} → ${newStatus}`);
        return {
          success: true,
          message: `ステータスを「${newStatus}」に更新しました`
        };
      } else {
        return {
          success: false,
          message: result.error || 'ステータス更新に失敗しました'
        };
      }

    } catch (e) {
      logError('PatientService.updateStatus', e);
      return {
        success: false,
        message: 'システムエラー: ' + e.message
      };
    }
  },

  /**
   * 出力日時を記録
   * @param {string} patientId - 受診者ID
   * @returns {Object} 結果
   */
  recordExport(patientId) {
    try {
      const result = PatientDAO.updateFields(patientId, {
        exportDate: new Date(),
        status: '出力済'
      });

      if (result.success) {
        logInfo(`PatientService.recordExport: ${patientId}`);
        return { success: true, message: '出力日時を記録しました' };
      } else {
        return { success: false, message: result.error };
      }

    } catch (e) {
      logError('PatientService.recordExport', e);
      return { success: false, message: 'システムエラー: ' + e.message };
    }
  },

  // ============================================
  // 統計・集計
  // ============================================

  /**
   * ステータス別件数を取得
   * @returns {Object} ステータス別件数
   */
  getStatusCounts() {
    try {
      const all = PatientDAO.getAll();
      const counts = {
        '入力中': 0,
        '入力完了': 0,
        '確認済': 0,
        '出力済': 0,
        total: all.length
      };

      for (const entity of all) {
        const status = entity.status || '入力中';
        if (counts.hasOwnProperty(status)) {
          counts[status]++;
        }
      }

      return counts;
    } catch (e) {
      logError('PatientService.getStatusCounts', e);
      return { total: 0 };
    }
  },

  /**
   * 総件数を取得
   * @returns {number} 件数
   */
  getCount() {
    return PatientDAO.count();
  }
};
