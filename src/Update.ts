class Update {
  private static readonly HANDLER = 'update';
  private static readonly STEP_PROPERTY = 'UPDATE_STEP';
  private static readonly GAMES_STEP = 'games';
  private static readonly TITLES_STEP = 'titles';
  private static readonly TRIGGER_INTERVAL_MINUTES = 5;
  private static readonly LEGACY_HANDLERS = [
    'updateGames',
    'updateRankings',
    'updateTitles',
    'updateRatings',
  ];

  static run(isScheduledRun: boolean): void {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(0)) {
      Logger.log('Skipped update because another update is already running.');
      return;
    }
    try {
      if (isScheduledRun) {
        Update.runScheduled();
      } else {
        Update.start();
      }
    } finally {
      lock.releaseLock();
    }
  }

  private static start(): void {
    const properties = PropertiesService.getScriptProperties();
    // Retire independent triggers from earlier deployments before creating the shared queue.
    Update.LEGACY_HANDLERS.forEach((handler) => Triggers.deleteAll(handler));
    Triggers.deleteAll(Update.HANDLER);
    properties.deleteProperty(Update.STEP_PROPERTY);

    UpdateRankings.run();
    UpdateRatings.run();

    // The phase is persisted so one recurring trigger can finish one batch type before starting the next.
    properties.setProperty(Update.STEP_PROPERTY, Update.GAMES_STEP);
    Triggers.ensure(Update.HANDLER, Update.TRIGGER_INTERVAL_MINUTES);
  }

  private static runScheduled(): void {
    const properties = PropertiesService.getScriptProperties();
    const step = properties.getProperty(Update.STEP_PROPERTY);
    if (step === Update.GAMES_STEP) {
      if (!UpdateGames.run()) {
        properties.setProperty(Update.STEP_PROPERTY, Update.TITLES_STEP);
      }
      return;
    }
    if (step === Update.TITLES_STEP) {
      if (!UpdateTitles.run()) {
        properties.deleteProperty(Update.STEP_PROPERTY);
        Triggers.deleteAll(Update.HANDLER);
      }
      return;
    }

    properties.deleteProperty(Update.STEP_PROPERTY);
    Triggers.deleteAll(Update.HANDLER);
  }
}

function update(event?: unknown): void {
  Update.run(event !== undefined);
}
