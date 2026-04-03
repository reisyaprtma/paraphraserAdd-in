/*
  Warnings:

  - You are about to drop the column `bleuScore` on the `ParaphraseLog` table. All the data in the column will be lost.
  - You are about to drop the column `paraPluieScore` on the `ParaphraseLog` table. All the data in the column will be lost.
  - You are about to drop the column `paraphrase_completionTokens` on the `ParaphraseLog` table. All the data in the column will be lost.
  - You are about to drop the column `paraphrase_latency` on the `ParaphraseLog` table. All the data in the column will be lost.
  - You are about to drop the column `paraphrase_promptTokens` on the `ParaphraseLog` table. All the data in the column will be lost.
  - You are about to drop the column `paraphrase_totalTokens` on the `ParaphraseLog` table. All the data in the column will be lost.
  - You are about to drop the column `parapluie_completionTokens` on the `ParaphraseLog` table. All the data in the column will be lost.
  - You are about to drop the column `parapluie_latency` on the `ParaphraseLog` table. All the data in the column will be lost.
  - You are about to drop the column `parapluie_promptTokens` on the `ParaphraseLog` table. All the data in the column will be lost.
  - You are about to drop the column `parapluie_totalTokens` on the `ParaphraseLog` table. All the data in the column will be lost.

*/
-- AlterTable
ALTER TABLE "ParaphraseLog" DROP COLUMN "bleuScore",
DROP COLUMN "paraPluieScore",
DROP COLUMN "paraphrase_completionTokens",
DROP COLUMN "paraphrase_latency",
DROP COLUMN "paraphrase_promptTokens",
DROP COLUMN "paraphrase_totalTokens",
DROP COLUMN "parapluie_completionTokens",
DROP COLUMN "parapluie_latency",
DROP COLUMN "parapluie_promptTokens",
DROP COLUMN "parapluie_totalTokens";

-- CreateTable
CREATE TABLE "Paraphrase" (
    "id" TEXT NOT NULL,
    "sourceText" TEXT NOT NULL,
    "tokenizedInput" TEXT,
    "paraphrasedText" TEXT NOT NULL,
    "iteration" INTEGER NOT NULL DEFAULT 1,
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

    CONSTRAINT "Paraphrase_pkey" PRIMARY KEY ("id")
);
