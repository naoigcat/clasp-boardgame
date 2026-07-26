/**
 * Coordinates synchronous and batched spreadsheet updates through one trigger.
 *
 * Rankings and Ratings usually fit in the menu execution. Games and Titles are
 * paced across later trigger runs so UrlFetch delays and Apps Script time limits
 * do not abort a full refresh. All paths share one lock and one handler name.
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
      // The active invocation owns the next batch schedule, so waiting here
      // would only spend Apps Script execution time while it owns the sheets.
      Logger.log('Skipped update because another update is already running.');
      return;
    }

    try {
      if (isScheduledExecution) {
        UpdateCoordinator.resumePendingPhase();
      } else {
        // A menu click always restarts: prior queue state and its trigger are
        // discarded so users get a fresh Rankings/Ratings snapshot immediately.
        UpdateCoordinator.startNewUpdate();
      }
    } finally {
      lock.releaseLock();
    }
  }

  /**
   * Performs work that fits in one execution and schedules the first batch.
   *
   * Fast sheet replacements run first; the Games phase is persisted before the
   * trigger is created so a trigger that fires early still knows what to do.
   */
  private static startNewUpdate(): void {
    UpdateCoordinator.removeSupersededTriggers();
    ScriptPropertyStore.remove(UPDATE_QUEUE_CONFIG.STEP_PROPERTY_KEY);

    // These imports replace complete sheet snapshots and normally fit in one
    // execution; the slower per-game work is deliberately deferred to batches.
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

    // A missing or obsolete state must not leave an orphaned trigger running
    // forever after properties have been edited or a deployment has changed.
    UpdateCoordinator.finishUpdate();
  }

  /**
   * Continues BoardGameGeek batches and switches to titles when they finish.
   *
   * `GameUpdater.run` returns true while stale rows remain so the same phase
   * key is reused. Switching only after false avoids starting Titles early.
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
   *
   * Without this cleanup the five-minute trigger would keep firing with no
   * useful phase and waste project quota.
   */
  private static finishUpdate(): void {
    ScriptPropertyStore.remove(UPDATE_QUEUE_CONFIG.STEP_PROPERTY_KEY);
    TriggerManager.removeAll(UPDATE_QUEUE_CONFIG.HANDLER_NAME);
  }

  /**
   * Retires old per-service triggers before a new unified queue is created.
   *
   * Also clears the current unified handler so a restart does not stack a
   * second schedule on top of the previous cycle.
   */
  private static removeSupersededTriggers(): void {
    LEGACY_UPDATE_HANDLER_NAMES.forEach((handlerName) => {
      TriggerManager.removeAll(handlerName);
    });
    TriggerManager.removeAll(UPDATE_QUEUE_CONFIG.HANDLER_NAME);
  }
}
