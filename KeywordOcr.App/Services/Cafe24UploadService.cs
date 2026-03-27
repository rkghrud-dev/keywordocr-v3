using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace KeywordOcr.App.Services;

public sealed class Cafe24UploadService
{
    private readonly Cafe24ConfigStore _configStore;
    private readonly Cafe24ApiClient _apiClient = new();

    public Cafe24UploadService(string v2Root, string legacyRoot)
    {
        _configStore = new Cafe24ConfigStore(v2Root, legacyRoot);
    }

    public async Task<Cafe24UploadResult> UploadAsync(
        string sourcePath,
        string exportRoot,
        Cafe24UploadOptions? overrideOptions = null,
        IProgress<string>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var options = overrideOptions ?? _configStore.LoadUploadOptions(exportRoot);
        if (string.IsNullOrWhiteSpace(options.ExportDir))
        {
            options.ExportDir = string.IsNullOrWhiteSpace(exportRoot) ? @"C:\code\exports" : exportRoot;
        }

        var tokenState = _configStore.LoadTokenState(options.TokenFilePath);
        ValidateTokenConfig(tokenState.Config);
        var workingDirectory = Cafe24UploadSupport.ResolveWorkingDirectory(sourcePath, exportRoot, options);
        progress?.Report($"작업 폴더: {workingDirectory}");

        var uploadWorkbookPath = Cafe24UploadSupport.FindLatestFileInDirectory(workingDirectory, "업로드용_*.xlsx");
        if (uploadWorkbookPath is null)
        {
            throw new FileNotFoundException("업로드용 엑셀을 찾지 못했습니다. 먼저 업로드용 엑셀을 생성해 주세요.", workingDirectory);
        }

        // sourcePath가 LLM 결과 파일이면 키워드 데이터를 거기서 읽기
        var keywordSourcePath = File.Exists(sourcePath) && sourcePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
            ? sourcePath : uploadWorkbookPath;
        var uploadNames = Cafe24UploadSupport.ReadUploadProductNames(uploadWorkbookPath);
        var keywordData = Cafe24UploadSupport.ReadProductKeywordData(keywordSourcePath);
        var gptWorkbookPath = Cafe24UploadSupport.FindLatestFileInDirectory(workingDirectory, "상품전처리GPT_*.xlsx");
        var optionPriceMap = Cafe24UploadSupport.LoadOptionSupplyPrices(gptWorkbookPath ?? uploadWorkbookPath);
        var priceReview = Cafe24UploadSupport.LoadPriceReview(options.PriceDataPath);
        var dateTag = Cafe24UploadSupport.ExtractDateTag(uploadWorkbookPath) ?? options.DateTag ?? DateTime.Now.ToString("yyyyMMdd", CultureInfo.InvariantCulture);
        var imageRoot = Cafe24UploadSupport.ResolveImageRoot(workingDirectory, options, dateTag);
        var folders = Cafe24UploadSupport.GetGsFolders(imageRoot);
        if (folders.Count == 0)
        {
            throw new DirectoryNotFoundException("GS 이미지 폴더를 찾지 못했습니다.");
        }

        if (priceReview.CheckedGs.Count == 0 && !string.IsNullOrWhiteSpace(options.GsListPath) && File.Exists(options.GsListPath))
        {
            var wanted = File.ReadAllLines(options.GsListPath)
                .Select(line => line.Trim().ToUpperInvariant())
                .Where(line => !string.IsNullOrWhiteSpace(line))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            folders = folders.Where(folder => wanted.Contains(folder.Name.ToUpperInvariant())).ToList();
        }

        progress?.Report($"업로드 엑셀: {Path.GetFileName(uploadWorkbookPath)}");
        progress?.Report($"이미지 폴더: {imageRoot}");
        progress?.Report($"대상 상품: {folders.Count}개");

        var products = await ExecuteWithRefreshAsync(tokenState, cfg => _apiClient.GetSellingProductsAsync(cfg, cancellationToken), cancellationToken);
        if (products.Count == 0)
        {
            throw new InvalidDataException("Cafe24 판매중 상품을 찾지 못했습니다.");
        }
        progress?.Report($"Cafe24 판매중 상품 로드: {products.Count}개");

        var productsByName = products
            .GroupBy(product => product.ProductName, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);

        var successCount = 0;
        var errorCount = 0;
        var skippedCount = 0;
        var logRows = new List<Dictionary<string, string>>();

        for (var index = 0; index < folders.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var folder = folders[index];
            var gs = folder.Name.ToUpperInvariant();
            var gs9 = gs.Length >= 9 ? gs[..9] : gs;
            progress?.Report($"[{index + 1}/{folders.Count}] {gs} 처리 시작");

            var selection = priceReview.ImageSelections.TryGetValue(gs9, out var imageSelection) ? imageSelection : null;
            var (mainImagePath, additionalImagePaths) = selection is null
                ? Cafe24UploadSupport.PickImages(folder.FullName, options.MainIndex, options.AddStart, options.AddMax)
                : Cafe24UploadSupport.PickImagesBySelection(folder.FullName, selection);

            if (string.IsNullOrWhiteSpace(mainImagePath))
            {
                skippedCount += 1;
                logRows.Add(Cafe24UploadSupport.CreateLogRow(gs, status: "NO_MAIN_IMAGE"));
                progress?.Report($"[{index + 1}/{folders.Count}] {gs} 대표이미지 없음");
                continue;
            }

            var matchedProduct = Cafe24UploadSupport.FindMatchingProduct(gs, uploadNames, products, productsByName, options.MatchMode, options.MatchPrefix);
            if (matchedProduct is null)
            {
                skippedCount += 1;
                logRows.Add(Cafe24UploadSupport.CreateLogRow(gs, status: "NO_PRODUCT_MATCH", mainImagePath: mainImagePath, additionalImagePaths: additionalImagePaths, selection: selection));
                progress?.Report($"[{index + 1}/{folders.Count}] {gs} 상품 매칭 실패");
                continue;
            }

            var uploadSucceeded = false;
            var uploadError = string.Empty;
            for (var attempt = 0; attempt <= options.RetryCount; attempt++)
            {
                try
                {
                    await ExecuteWithRefreshAsync(tokenState, cfg => _apiClient.UploadMainImageAsync(cfg, matchedProduct.ProductNo, mainImagePath, cancellationToken), cancellationToken);
                    await UploadAdditionalImagesWithRecoveryAsync(
                        tokenState,
                        matchedProduct.ProductNo,
                        additionalImagePaths,
                        progress,
                        $"[{index + 1}/{folders.Count}] {gs}",
                        cancellationToken);
                    uploadSucceeded = true;
                    break;
                }
                catch (Cafe24ReauthenticationRequiredException)
                {
                    throw;
                }
                catch (Exception ex) when (attempt < options.RetryCount)
                {
                    uploadError = Cafe24UploadSupport.UnwrapMessage(ex);
                    progress?.Report($"[{index + 1}/{folders.Count}] {gs} 업로드 재시도 {attempt + 1}/{options.RetryCount}");
                    await Task.Delay(TimeSpan.FromSeconds(options.RetryDelaySeconds), cancellationToken);
                }
                catch (Exception ex)
                {
                    uploadError = Cafe24UploadSupport.UnwrapMessage(ex);
                    break;
                }
            }

            if (!uploadSucceeded)
            {
                errorCount += 1;
                logRows.Add(Cafe24UploadSupport.CreateLogRow(
                    gs,
                    productNo: matchedProduct.ProductNo.ToString(CultureInfo.InvariantCulture),
                    status: "ERROR",
                    mainImagePath: mainImagePath,
                    additionalImagePaths: additionalImagePaths,
                    selection: selection,
                    error: uploadError));
                progress?.Report($"[{index + 1}/{folders.Count}] {gs} 업로드 실패: {uploadError}");
                continue;
            }

            optionPriceMap.TryGetValue(gs9, out var optionData);
            optionData ??= new List<OptionSupplyItem>();
            var priceStatus = await ApplyPriceAsync(tokenState, matchedProduct.ProductNo, gs9, optionData, priceReview, cancellationToken);
            if (HasPriceFailure(priceStatus))
            {
                errorCount += 1;
                logRows.Add(Cafe24UploadSupport.CreateLogRow(
                    gs,
                    productNo: matchedProduct.ProductNo.ToString(CultureInfo.InvariantCulture),
                    status: "PARTIAL",
                    mainImagePath: mainImagePath,
                    additionalImagePaths: additionalImagePaths,
                    selection: selection,
                    priceStatus: priceStatus));
                progress?.Report($"[{index + 1}/{folders.Count}] {gs} 이미지 업로드 완료, 가격 적용 실패 ({priceStatus})");
                continue;
            }

            // ── 상품명 / 검색어설정 / 검색키워드 업데이트 ──
            var nameStatus = "";
            if (keywordData.TryGetValue(gs9, out var kwInfo) && !string.IsNullOrWhiteSpace(kwInfo.ProductName))
            {
                try
                {
                    // 상품명에서 GS코드 제거 (이미 제거된 경우도 안전)
                    var cleanName = System.Text.RegularExpressions.Regex.Replace(kwInfo.ProductName, @"GS\d{7,9}[A-Z]?\s*", "").Trim();
                    if (!string.IsNullOrWhiteSpace(cleanName))
                    {
                        await ExecuteWithRefreshAsync(tokenState, cfg =>
                            _apiClient.UpdateProductAsync(cfg, matchedProduct.ProductNo, cleanName, kwInfo.ProductTag, kwInfo.SearchKeyword, cancellationToken), cancellationToken);
                        nameStatus = "NAME_OK";
                        progress?.Report($"[{index + 1}/{folders.Count}] {gs} 상품명+검색어 업데이트 완료");
                    }
                }
                catch (Exception nameEx)
                {
                    nameStatus = $"NAME_ERROR: {Cafe24UploadSupport.UnwrapMessage(nameEx)}";
                    progress?.Report($"[{index + 1}/{folders.Count}] {gs} 상품명 업데이트 실패: {nameStatus}");
                }
            }

            successCount += 1;
            logRows.Add(Cafe24UploadSupport.CreateLogRow(
                gs,
                productNo: matchedProduct.ProductNo.ToString(CultureInfo.InvariantCulture),
                status: "OK",
                mainImagePath: mainImagePath,
                additionalImagePaths: additionalImagePaths,
                selection: selection,
                priceStatus: string.IsNullOrEmpty(nameStatus) ? priceStatus : $"{priceStatus}|{nameStatus}"));
            progress?.Report($"[{index + 1}/{folders.Count}] {gs} 업로드 완료 ({priceStatus}{(string.IsNullOrEmpty(nameStatus) ? "" : $"|{nameStatus}")})");

        }

        var logPath = Cafe24UploadSupport.WriteLogWorkbook(logRows, workingDirectory, options.LogPath);
        progress?.Report($"업로드 로그 저장: {logPath}");

        return new Cafe24UploadResult(workingDirectory, uploadWorkbookPath, logPath, folders.Count, successCount, errorCount, skippedCount);
    }

    public async Task<Cafe24UploadResult> UploadBMarketAsync(
        string sourcePath,
        string exportRoot,
        Cafe24UploadOptions? overrideOptions = null,
        IProgress<string>? progress = null,
        CancellationToken cancellationToken = default)
    {
        Cafe24TokenState tokenState;
        try
        {
            tokenState = _configStore.LoadTokenStateB();
        }
        catch (FileNotFoundException ex)
        {
            progress?.Report($"B마켓 업로드 스킵: {ex.Message}");
            return new Cafe24UploadResult("", "", "", 0, 0, 0, 0);
        }
        ValidateTokenConfig(tokenState.Config);

        var options = overrideOptions ?? _configStore.LoadUploadOptions(exportRoot);
        if (string.IsNullOrWhiteSpace(options.ExportDir))
            options.ExportDir = string.IsNullOrWhiteSpace(exportRoot) ? @"C:\code\exports" : exportRoot;

        var workingDirectory = Cafe24UploadSupport.ResolveWorkingDirectory(sourcePath, exportRoot, options);
        progress?.Report($"[B마켓] 작업 폴더: {workingDirectory}");

        var uploadWorkbookPath = Cafe24UploadSupport.FindLatestFileInDirectory(workingDirectory, "업로드용_*.xlsx");
        if (uploadWorkbookPath is null)
        {
            progress?.Report("[B마켓] 업로드용 엑셀을 찾지 못했습니다.");
            return new Cafe24UploadResult(workingDirectory, "", "", 0, 0, 0, 0);
        }

        var keywordSourcePath = File.Exists(sourcePath) && sourcePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
            ? sourcePath : uploadWorkbookPath;
        var keywordData = Cafe24UploadSupport.ReadProductKeywordData(keywordSourcePath, "B마켓");
        var uploadNames = Cafe24UploadSupport.ReadUploadProductNames(uploadWorkbookPath);
        var gptWorkbookPath = Cafe24UploadSupport.FindLatestFileInDirectory(workingDirectory, "상품전처리GPT_*.xlsx");
        var optionPriceMap = Cafe24UploadSupport.LoadOptionSupplyPrices(gptWorkbookPath ?? uploadWorkbookPath);
        var priceReview = Cafe24UploadSupport.LoadPriceReview(options.PriceDataPath);
        var dateTag = Cafe24UploadSupport.ExtractDateTag(uploadWorkbookPath) ?? options.DateTag ?? DateTime.Now.ToString("yyyyMMdd", CultureInfo.InvariantCulture);

        // B마켓은 listing_images_B/ 폴더 사용
        var imageRootBase = Cafe24UploadSupport.ResolveImageRoot(workingDirectory, options, dateTag);
        var imageRootB = imageRootBase.Replace("listing_images", "listing_images_B");
        if (!Directory.Exists(imageRootB))
            imageRootB = imageRootBase; // fallback to A market images

        var folders = Cafe24UploadSupport.GetGsFolders(imageRootB);
        if (folders.Count == 0)
        {
            progress?.Report("[B마켓] GS 이미지 폴더를 찾지 못했습니다.");
            return new Cafe24UploadResult(workingDirectory, uploadWorkbookPath, "", 0, 0, 0, 0);
        }

        progress?.Report($"[B마켓] 이미지 폴더: {imageRootB}");
        progress?.Report($"[B마켓] 대상 상품: {folders.Count}개");

        var products = await ExecuteWithRefreshAsync(tokenState, cfg => _apiClient.GetSellingProductsAsync(cfg, cancellationToken), cancellationToken);
        progress?.Report($"[B마켓] Cafe24 판매중 상품 로드: {products.Count}개");

        var productsByName = products
            .GroupBy(product => product.ProductName, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);

        var successCount = 0;
        var errorCount = 0;
        var skippedCount = 0;
        var logRows = new List<Dictionary<string, string>>();

        for (var index = 0; index < folders.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var folder = folders[index];
            var gs = folder.Name.ToUpperInvariant();
            var gs9 = gs.Length >= 9 ? gs[..9] : gs;
            progress?.Report($"[B마켓] [{index + 1}/{folders.Count}] {gs} 처리 시작");

            var (mainImagePath, additionalImagePaths) = Cafe24UploadSupport.PickImages(folder.FullName, options.MainIndex, options.AddStart, options.AddMax);
            if (string.IsNullOrWhiteSpace(mainImagePath))
            {
                skippedCount += 1;
                continue;
            }

            var matchedProduct = Cafe24UploadSupport.FindMatchingProduct(gs, uploadNames, products, productsByName, options.MatchMode, options.MatchPrefix);
            if (matchedProduct is null)
            {
                skippedCount += 1;
                progress?.Report($"[B마켓] [{index + 1}/{folders.Count}] {gs} 상품 매칭 실패");
                continue;
            }

            var uploadSucceeded = false;
            var uploadError = string.Empty;
            for (var attempt = 0; attempt <= options.RetryCount; attempt++)
            {
                try
                {
                    await ExecuteWithRefreshAsync(tokenState, cfg => _apiClient.UploadMainImageAsync(cfg, matchedProduct.ProductNo, mainImagePath, cancellationToken), cancellationToken);
                    await UploadAdditionalImagesWithRecoveryAsync(tokenState, matchedProduct.ProductNo, additionalImagePaths, progress, $"[B마켓] [{index + 1}/{folders.Count}] {gs}", cancellationToken);
                    uploadSucceeded = true;
                    break;
                }
                catch (Cafe24ReauthenticationRequiredException) { throw; }
                catch (Exception ex) when (attempt < options.RetryCount)
                {
                    uploadError = Cafe24UploadSupport.UnwrapMessage(ex);
                    await Task.Delay(TimeSpan.FromSeconds(options.RetryDelaySeconds), cancellationToken);
                }
                catch (Exception ex) { uploadError = Cafe24UploadSupport.UnwrapMessage(ex); break; }
            }

            if (!uploadSucceeded) { errorCount += 1; continue; }

            // 상품명 / 검색어 업데이트 (B마켓 시트 데이터)
            if (keywordData.TryGetValue(gs9, out var kwInfo) && !string.IsNullOrWhiteSpace(kwInfo.ProductName))
            {
                try
                {
                    var cleanName = System.Text.RegularExpressions.Regex.Replace(kwInfo.ProductName, @"GS\d{7,9}[A-Z]?\s*", "").Trim();
                    if (!string.IsNullOrWhiteSpace(cleanName))
                    {
                        await ExecuteWithRefreshAsync(tokenState, cfg =>
                            _apiClient.UpdateProductAsync(cfg, matchedProduct.ProductNo, cleanName, kwInfo.ProductTag, kwInfo.SearchKeyword, cancellationToken), cancellationToken);
                    }
                }
                catch (Exception nameEx)
                {
                    progress?.Report($"[B마켓] [{index + 1}/{folders.Count}] {gs} 상품명 업데이트 실패: {Cafe24UploadSupport.UnwrapMessage(nameEx)}");
                }
            }

            successCount += 1;
            progress?.Report($"[B마켓] [{index + 1}/{folders.Count}] {gs} 업로드 완료");
        }

        var logPath = Cafe24UploadSupport.WriteLogWorkbook(logRows, workingDirectory, null);
        progress?.Report($"[B마켓] 업로드 완료: 성공 {successCount} / 오류 {errorCount} / 스킵 {skippedCount}");

        return new Cafe24UploadResult(workingDirectory, uploadWorkbookPath, logPath, folders.Count, successCount, errorCount, skippedCount);
    }

    private async Task UploadAdditionalImagesWithRecoveryAsync(
        Cafe24TokenState tokenState,
        int productNo,
        IReadOnlyList<string> additionalImagePaths,
        IProgress<string>? progress,
        string progressPrefix,
        CancellationToken cancellationToken)
    {
        if (additionalImagePaths.Count == 0)
        {
            return;
        }

        try
        {
            foreach (var imagePath in additionalImagePaths)
            {
                await ExecuteWithRefreshAsync(tokenState, cfg => _apiClient.UploadAdditionalImageAsync(cfg, productNo, imagePath, cancellationToken), cancellationToken);
            }
        }
        catch (Exception ex) when (IsAdditionalImageLimitError(ex))
        {
            progress?.Report($"{progressPrefix} 기존 추가이미지 한도 도달로 전체 삭제 후 재업로드");
            await ExecuteWithRefreshAsync(tokenState, cfg => _apiClient.DeleteAdditionalImagesAsync(cfg, productNo, cancellationToken), cancellationToken);

            foreach (var imagePath in additionalImagePaths)
            {
                await ExecuteWithRefreshAsync(tokenState, cfg => _apiClient.UploadAdditionalImageAsync(cfg, productNo, imagePath, cancellationToken), cancellationToken);
            }
        }
    }

    private async Task<string> ApplyPriceAsync(
        Cafe24TokenState tokenState,
        int productNo,
        string gs9,
        IReadOnlyList<OptionSupplyItem> optionData,
        PriceReviewData priceReview,
        CancellationToken cancellationToken)
    {
        var isOptionProduct = priceReview.CheckedGs.Contains(gs9) || optionData.Count > 1;
        if (!isOptionProduct)
        {
            return "PRICE_SKIP_SINGLE";
        }

        if (priceReview.EditedAmounts.TryGetValue(gs9, out var editedAmounts) && editedAmounts.Count > 0)
        {
            return await ApplyEditedAmountsAsync(tokenState, productNo, optionData, editedAmounts, cancellationToken);
        }

        if (optionData.Count <= 1)
        {
            return "PRICE_SKIP";
        }

        return await ApplyVariantPricesAsync(tokenState, productNo, optionData, cancellationToken);
    }

    private async Task<string> ApplyEditedAmountsAsync(
        Cafe24TokenState tokenState,
        int productNo,
        IReadOnlyList<OptionSupplyItem> optionData,
        IReadOnlyList<decimal> editedAmounts,
        CancellationToken cancellationToken)
    {
        var variants = await ExecuteWithRefreshAsync(tokenState, cfg => _apiClient.GetVariantsAsync(cfg, productNo, cancellationToken), cancellationToken);
        if (variants.Count == editedAmounts.Count)
        {
            var assignments = variants
                .Select((variant, index) => (Variant: variant, AdditionalAmount: editedAmounts[index]))
                .ToList();
            return await UpdateVariantAmountsAsync(tokenState, productNo, variants, assignments, cancellationToken);
        }

        if (optionData.Count == editedAmounts.Count)
        {
            var desired = optionData
                .Select((item, index) => (Suffix: item.Suffix, AdditionalAmount: editedAmounts[index]))
                .ToList();
            if (TryBuildVariantAssignments(variants, desired, out var assignments, out var matchMessage))
            {
                return await UpdateVariantAmountsAsync(tokenState, productNo, variants, assignments, cancellationToken);
            }

            return $"PRICE_ERROR: variant수({variants.Count})≠편집금액수({editedAmounts.Count}) / {matchMessage}";
        }

        return $"PRICE_ERROR: variant수({variants.Count})≠편집금액수({editedAmounts.Count})";
    }

    private async Task<string> ApplyVariantPricesAsync(Cafe24TokenState tokenState, int productNo, IReadOnlyList<OptionSupplyItem> optionData, CancellationToken cancellationToken)
    {
        var supplyPrices = optionData.Select(item => item.SupplyPrice).ToList();
        if (supplyPrices.Count > 1 && supplyPrices.Distinct().Count() <= 1)
        {
            return "PRICE_SKIP_SAME";
        }

        var calculation = Cafe24UploadSupport.CalcOptionPrices(supplyPrices);
        if (calculation.AdditionalAmounts.All(amount => amount == 0))
        {
            return "PRICE_SKIP_NO_DIFF";
        }

        var variants = await ExecuteWithRefreshAsync(tokenState, cfg => _apiClient.GetVariantsAsync(cfg, productNo, cancellationToken), cancellationToken);
        if (variants.Count != optionData.Count)
        {
            var desired = optionData
                .Select((item, index) => (Suffix: item.Suffix, AdditionalAmount: calculation.AdditionalAmounts[index]))
                .ToList();
            if (TryBuildVariantAssignments(variants, desired, out var assignments, out var matchMessage))
            {
                return await UpdateVariantAmountsAsync(tokenState, productNo, variants, assignments, cancellationToken);
            }

            return $"PRICE_ERROR: variant수({variants.Count})≠옵션수({optionData.Count}) / {matchMessage}";
        }

        var fullAssignments = variants
            .Select((variant, index) => (Variant: variant, AdditionalAmount: calculation.AdditionalAmounts[index]))
            .ToList();
        return await UpdateVariantAmountsAsync(tokenState, productNo, variants, fullAssignments, cancellationToken);
    }

    private async Task<string> UpdateVariantAmountsAsync(
        Cafe24TokenState tokenState,
        int productNo,
        IReadOnlyList<Cafe24Variant> allVariants,
        IReadOnlyList<(Cafe24Variant Variant, decimal AdditionalAmount)> assignments,
        CancellationToken cancellationToken)
    {
        var errors = new List<string>();
        foreach (var assignment in assignments)
        {
            try
            {
                await ExecuteWithRefreshAsync(tokenState, cfg => _apiClient.UpdateVariantAsync(cfg, productNo, assignment.Variant.VariantCode, assignment.AdditionalAmount, cancellationToken), cancellationToken);
            }
            catch (Cafe24ReauthenticationRequiredException)
            {
                throw;
            }
            catch (Exception ex)
            {
                var label = string.IsNullOrWhiteSpace(assignment.Variant.OptionSummary)
                    ? assignment.Variant.VariantCode
                    : $"{assignment.Variant.VariantCode}:{assignment.Variant.OptionSummary}";
                errors.Add($"{label}:{Cafe24UploadSupport.UnwrapMessage(ex)}");
            }
        }

        if (errors.Count > 0)
        {
            return $"PRICE_PARTIAL: {string.Join("; ", errors)}";
        }

        return assignments.Count == allVariants.Count
            ? "PRICE_OK"
            : $"PRICE_OK_MATCHED: {assignments.Count}/{allVariants.Count}";
    }

    private static bool TryBuildVariantAssignments(
        IReadOnlyList<Cafe24Variant> variants,
        IReadOnlyList<(string Suffix, decimal AdditionalAmount)> desired,
        out List<(Cafe24Variant Variant, decimal AdditionalAmount)> assignments,
        out string matchMessage)
    {
        assignments = new List<(Cafe24Variant Variant, decimal AdditionalAmount)>();
        var usedIndices = new HashSet<int>();

        foreach (var item in desired)
        {
            var normalizedSuffix = NormalizeOptionText(item.Suffix);
            if (string.IsNullOrWhiteSpace(normalizedSuffix))
            {
                matchMessage = $"옵션매칭실패: 옵션명 없음 ({DescribeVariants(variants)})";
                assignments.Clear();
                return false;
            }

            var candidates = variants
                .Select((variant, index) => new { Variant = variant, Index = index, Score = GetMatchScore(variant, normalizedSuffix) })
                .Where(candidate => candidate.Score > 0 && !usedIndices.Contains(candidate.Index))
                .OrderByDescending(candidate => candidate.Score)
                .ToList();
            if (candidates.Count == 0)
            {
                matchMessage = $"옵션매칭실패: {item.Suffix} -> {DescribeVariants(variants)}";
                assignments.Clear();
                return false;
            }

            var topScore = candidates[0].Score;
            var bestMatches = candidates.Where(candidate => candidate.Score == topScore).ToList();
            if (bestMatches.Count != 1)
            {
                matchMessage = $"옵션매칭중복: {item.Suffix} -> {string.Join(", ", bestMatches.Select(candidate => DescribeVariant(candidate.Variant)))}";
                assignments.Clear();
                return false;
            }

            usedIndices.Add(bestMatches[0].Index);
            assignments.Add((bestMatches[0].Variant, item.AdditionalAmount));
        }

        matchMessage = assignments.Count == variants.Count
            ? "옵션 전체 매칭"
            : $"옵션 부분 매칭 {assignments.Count}/{variants.Count}";
        return assignments.Count > 0;
    }

    private static int GetMatchScore(Cafe24Variant variant, string normalizedSuffix)
    {
        var best = 0;
        foreach (var optionValue in variant.OptionValues)
        {
            var normalizedOption = NormalizeOptionText(optionValue);
            if (string.IsNullOrWhiteSpace(normalizedOption))
            {
                continue;
            }

            if (string.Equals(normalizedOption, normalizedSuffix, StringComparison.OrdinalIgnoreCase))
            {
                return 400;
            }
            if (normalizedOption.EndsWith(normalizedSuffix, StringComparison.OrdinalIgnoreCase))
            {
                best = Math.Max(best, 300);
                continue;
            }
            if (normalizedOption.Contains(normalizedSuffix, StringComparison.OrdinalIgnoreCase))
            {
                best = Math.Max(best, 200);
                continue;
            }
            if (normalizedSuffix.Contains(normalizedOption, StringComparison.OrdinalIgnoreCase))
            {
                best = Math.Max(best, 100);
            }
        }

        return best;
    }

    private static bool HasPriceFailure(string priceStatus)
    {
        return priceStatus.StartsWith("PRICE_ERROR", StringComparison.OrdinalIgnoreCase)
            || priceStatus.StartsWith("PRICE_PARTIAL", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsAdditionalImageLimitError(Exception exception)
    {
        var message = Cafe24UploadSupport.UnwrapMessage(exception);
        return message.Contains("number of additional image", StringComparison.OrdinalIgnoreCase)
            || message.Contains("additional image cannot exceeds 20", StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizeOptionText(string? value)
    {
        return Regex.Replace(value ?? string.Empty, @"[^0-9A-Za-z가-힣]", string.Empty).ToUpperInvariant();
    }

    private static string DescribeVariants(IReadOnlyList<Cafe24Variant> variants)
    {
        return string.Join(", ", variants.Select(DescribeVariant));
    }

    private static string DescribeVariant(Cafe24Variant variant)
    {
        return string.IsNullOrWhiteSpace(variant.OptionSummary)
            ? variant.VariantCode
            : variant.OptionSummary;
    }

    private async Task ExecuteWithRefreshAsync(Cafe24TokenState tokenState, Func<Cafe24TokenConfig, Task> action, CancellationToken cancellationToken)
    {
        await ExecuteWithRefreshAsync(tokenState, async config =>
        {
            await action(config);
            return true;
        }, cancellationToken);
    }

    private async Task<T> ExecuteWithRefreshAsync<T>(Cafe24TokenState tokenState, Func<Cafe24TokenConfig, Task<T>> action, CancellationToken cancellationToken)
    {
        try
        {
            return await action(tokenState.Config);
        }
        catch (Cafe24TokenExpiredException)
        {
            await _apiClient.RefreshAccessTokenAsync(tokenState.Config, cancellationToken);
            _configStore.SaveTokenConfig(tokenState.ConfigPath, tokenState.Config);
            return await action(tokenState.Config);
        }
    }

    private static void ValidateTokenConfig(Cafe24TokenConfig config)
    {
        if (string.IsNullOrWhiteSpace(config.MallId) || string.IsNullOrWhiteSpace(config.AccessToken))
        {
            throw new InvalidDataException("Cafe24 설정을 찾지 못했습니다. cafe24_token.txt에 MALL_ID와 ACCESS_TOKEN이 필요합니다.");
        }
    }
}
