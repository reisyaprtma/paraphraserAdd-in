-- CreateTable
CREATE TABLE "ParaphraseLog" (
    "id" TEXT NOT NULL,
    "sourceText" TEXT NOT NULL,
    "tokenizedInput" TEXT,
    "paraphrasedText" TEXT NOT NULL,
    "bleuScore" DOUBLE PRECISION,
    "paraPluieScore" DOUBLE PRECISION,
    "paraphrase_promptTokens" INTEGER,
    "paraphrase_completionTokens" INTEGER,
    "paraphrase_totalTokens" INTEGER,
    "paraphrase_latency" DOUBLE PRECISION,
    "parapluie_promptTokens" INTEGER,
    "parapluie_completionTokens" INTEGER,
    "parapluie_totalTokens" INTEGER,
    "parapluie_latency" DOUBLE PRECISION,
    "rating" INTEGER,
    "userComment" TEXT,
    "errorLog" TEXT,
    "iterations" INTEGER NOT NULL DEFAULT 1,
    "isSuccessful" BOOLEAN NOT NULL DEFAULT true,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "ParaphraseLog_pkey" PRIMARY KEY ("id")
);

-- CreateIndex
CREATE INDEX "ParaphraseLog_rating_idx" ON "ParaphraseLog"("rating");
