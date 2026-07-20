/**
 * Coordinates synchronous and batched spreadsheet updates through one trigger.
 */
class UpdateCoordinator {
  /**
   * Starts a manual update or advances the phase saved by a time-driven trigger.
   *
   * A script lock serializes invocations so that a manual click and a scheduled
   * run cannot mutate the same sheets at the same time.
   */
  static run(isScheduledExecution: boolean): void {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(0)) {
      Logger.log('Skipped update because another update is already running.');
      return;
    }

    try {
      if (isScheduledExecution) {
        UpdateCoordinator.resumePendingPhase();
      } else {
        UpdateCoordinator.startNewUpdate();
      }
    } finally {
      lock.releaseLock();
    }
  }

  /**
   * Performs work that fits in one execution and schedules the first batch.
   */
  private static startNewUpdate(): void {
    UpdateCoordinator.removeSupersededTriggers();
    ScriptPropertyStore.remove(UPDATE_QUEUE_CONFIG.STEP_PROPERTY_KEY);

    RankingUpdater.run();
    RatingUpdater.run();

    ScriptPropertyStore.set(
      UPDATE_QUEUE_CONFIG.STEP_PROPERTY_KEY,
      UPDATE_QUEUE_CONFIG.GAMES_STEP,
    );
    TriggerManager.ensureSingle(
      UPDATE_QUEUE_CONFIG.HANDLER_NAME,
      UPDATE_QUEUE_CONFIG.TRIGGER_INTERVAL_MINUTES,
    );
  }

  /**
   * Advances the phase recorded by the previous trigger invocation.
   */
  private static resumePendingPhase(): void {
    const nextStep = ScriptPropertyStore.getOptionalValue(
      UPDATE_QUEUE_CONFIG.STEP_PROPERTY_KEY,
    );

    if (nextStep === UPDATE_QUEUE_CONFIG.GAMES_STEP) {
      UpdateCoordinator.resumeGamesPhase();
      return;
    }

    if (nextStep === UPDATE_QUEUE_CONFIG.TITLES_STEP) {
      UpdateCoordinator.resumeTitlesPhase();
      return;
    }

    UpdateCoordinator.finishUpdate();
  }

  /**
   * Continues BoardGameGeek batches and switches to titles when they finish.
   */
  private static resumeGamesPhase(): void {
    if (!GameUpdater.run()) {
      ScriptPropertyStore.set(
        UPDATE_QUEUE_CONFIG.STEP_PROPERTY_KEY,
        UPDATE_QUEUE_CONFIG.TITLES_STEP,
      );
    }
  }

  /**
   * Continues title batches and removes queue state after the final batch.
   */
  private static resumeTitlesPhase(): void {
    if (!TitleUpdater.run()) {
      UpdateCoordinator.finishUpdate();
    }
  }

  /**
   * Removes transient state and the shared trigger after all work completes.
   */
  private static finishUpdate(): void {
    ScriptPropertyStore.remove(UPDATE_QUEUE_CONFIG.STEP_PROPERTY_KEY);
    TriggerManager.removeAll(UPDATE_QUEUE_CONFIG.HANDLER_NAME);
  }

  /**
   * Retires old per-service triggers before a new unified queue is created.
   */
  private static removeSupersededTriggers(): void {
    LEGACY_UPDATE_HANDLER_NAMES.forEach((handlerName) => {
      TriggerManager.removeAll(handlerName);
    });
    TriggerManager.removeAll(UPDATE_QUEUE_CONFIG.HANDLER_NAME);
  }
}
