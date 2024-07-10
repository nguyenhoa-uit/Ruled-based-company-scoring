from owlready2 import *
import math

# Create the ontology
onto = get_ontology("company_ontology.owl")

with onto:
    class Company(Thing):
        pass

    class FinancialRatio(Thing):
        pass

    class LiquidityRatio(FinancialRatio):
        pass

    class LeverageRatio(FinancialRatio):
        pass

    class EfficiencyRatio(FinancialRatio):
        pass

    class ProfitabilityRatio(FinancialRatio):
        pass

    # A-Liquidity ratios
    class CurrentRatio(LiquidityRatio):
        pass

    class QuickRatio(LiquidityRatio):
        pass

    class CashRatio(LiquidityRatio):
        pass

    # B-Leverage ratios
    class DebtRatio(LeverageRatio):
        pass

    class DebtToEquityRatio(LeverageRatio):
        pass

    class InterestCoverageRatio(LeverageRatio):
        pass

    # Efficiency ratios
    class InventoryTurnover(EfficiencyRatio):
        pass

    class AverageCollectionPeriod(EfficiencyRatio):
        pass

    class FixedAssetTurnover(EfficiencyRatio):
        pass

    # Profitability ratios
    class RevenueRatio(ProfitabilityRatio):
        pass

    class BasicProfitabilityRatio(ProfitabilityRatio):
        pass

    class ReturnOnAssets(ProfitabilityRatio):
        pass

# Function to calculate financial ratios
def calculate_ratios(company):
    # Liquidity ratios
    company.A1_current_ratio = company.short_term_assets / company.short_term_debt
    company.A2_quick_ratio = (company.short_term_assets - company.inventory) / company.short_term_debt
    company.A3_cash_ratio = company.cash / company.short_term_debt

    # Leverage ratios
    company.B1_debt_ratio = company.liabilities / company.total_assets
    company.B2_debt_to_equity_ratio = company.liabilities / company.equity
    company.B3_interest_coverage_ratio = company.ebit / company.interest_expenses

    # Efficiency ratios
    company.C1_inventory_turnover = company.cost_of_sales / company.average_inventory
    company.C2_average_collection_period = company.average_short_term_receivables / company.average_net_revenue
    company.C3_fixed_asset_turnover = company.net_revenue / company.average_net_fixed_assets

    # Profitability ratios
    company.D1_revenue_ratio = company.profit_after_tax / company.net_revenue
    company.D2_basic_profitability_ratio = company.ebit / company.average_total_assets
    company.D3_return_on_assets = company.profit_after_tax / company.average_total_assets

# Function to analyze company's overall status
def analyze_company(company):
    remarkA = ""
    if company.A1_current_ratio > 2 and company.A2_quick_ratio > 1 and company.A3_cash_ratio > 0.5:
        remarkA += "Công ty có khả năng thanh toán tốt."
    else:
        remarkA += "Công ty cần cải thiện khả năng thanh toán."
    remarkB = ""

    if company.B1_debt_ratio < 0.5 and company.B2_debt_to_equity_ratio < 1:
        remarkB += "Công ty có khả năng cân đối vốn tốt. "
    else:
        remarkB +=  "Mức độ cân đối vốn của công ty cần cải thiện. "

    if company.B3_interest_coverage_ratio > 3:
        remarkB +=  "Công ty có thể trả lãi vay dễ dàng. "
    else:
        remarkB += "Công ty cần cải thiện khả năng trả lãi vay. "
    
    remarkC = ""
    if company.C1_inventory_turnover > 5 and company.C2_average_collection_period < 60 and company.C3_fixed_asset_turnover > 2:
        remarkC += "Công ty hoạt động hiệu quả. "
    else:
        remarkC += "Công ty cần cải thiện hiệu quả hoạt động. "

    remarkD = ""
    if company.D1_revenue_ratio > 0.1 and company.D3_return_on_assets > 0.05 and company.D2_basic_profitability_ratio > 0.1:
        remarkD += "Công ty có khả năng sinh lợi tốt."
    else:
        remarkD += "Công ty cần cải thiện khả năng sinh lợi."
    return [remarkA, remarkB, remarkC, remarkD]

def get_score(value,index):
    heso_get_max=10
    sign_matrix=[True,True,True,False,False,True,True,False,True,True,True,True]
    medium_matrix=[2.0,1.0,0.5,0.5,1.0,3.0,5,60,2,0.1,0.05,0.1]
    sign=sign_matrix[index]
    medium=medium_matrix[index]
    if value<=0:
        return 0
    if sign:
        score=5+heso_get_max*math.log10(value/medium)/2
    else:
        score=5+heso_get_max*math.log10(medium/value)/2
    if score<0:
        score=0.0
    if score>10:
        score=10.0 
    return round(score,1)

def print_test(value):
    company = Company()
    company.total_assets = float(value["total_assets"])
    company.short_term_assets = float(value["short_term_assets"])
    company.cash = float(value["cash"])
    company.inventory = float(value["inventory"])
    company.long_term_assets = float(value["long_term_assets"])
    company.total_capital = float(value["total_capital"])

    company.liabilities = float(value["liabilities"])
    company.short_term_debt = float(value["short_term_debt"])
    company.short_term_receivables =float(value["short_term_receivables"])
    company.long_term_debt = float(value["long_term_debt"])
    company.equity = float(value["equity"])
    company.net_revenue =float(value["net_revenue"])
    company.cost_of_sales = float(value["cost_of_sales"])
    company.profit_before_tax = float(value["profit_before_tax"])
    company.interest_expenses = float(value["interest_expenses"])
    company.ebit =float(value["ebit"])
    company.profit_after_tax = float(value["profit_after_tax"])

    company.average_inventory = float(value["average_inventory"])
    company.average_short_term_receivables = float(value["average_short_term_receivables"])
    company.average_net_revenue = float(value["average_net_revenue"])
    company.average_net_fixed_assets = float(value["average_net_fixed_assets"])
    company.average_total_assets = float(value["average_total_assets"])

    # Calculate financial ratios
    calculate_ratios(company)

    # Print basic output

    remarkA, remarkB, remarkC, remarkD = analyze_company(company)
    s0=get_score(company.A1_current_ratio,0)
    s1=get_score(company.A2_quick_ratio,1)
    s2=get_score(company.A3_cash_ratio,2)

    s3=get_score(company.B1_debt_ratio,3)
    s4=get_score(company.B2_debt_to_equity_ratio,4)
    s5=get_score(company.B3_interest_coverage_ratio,5)

    s6=get_score(company.C1_inventory_turnover,6)
    s7=get_score(company.C2_average_collection_period,7)
    s8=get_score(company.C3_fixed_asset_turnover,8)
    s9=get_score(company.D1_revenue_ratio,9)
    s10=get_score(company.D2_basic_profitability_ratio,10)
    s11=get_score(company.D3_return_on_assets,11)

    comments=["","NHẬN XÉT VỀ TÌNH HÌNH TÀI CHÍNH CÔNG TY: TOTAL SCORE= {:.0f}/120.".format(s0+s1+s2+s3+s4+s5+s6+s7+s8+s9+s10+s11)]
    comments.append("A. Khả năng thanh toán: Score= {:.0f}/30.".format(s0+s1+s2))
    comments.append("A1: Tỷ số khả năng thanh toán hiện thời = {:.0f}%. Score= {:.0f}/10.".format(round(100*company.A1_current_ratio,1),s0))
    comments.append("A2: Tỷ số khả năng thanh toán nhanh ={:.0f}%. Score= {:.0f}/10.".format(round(100*company.A2_quick_ratio,1),s1))
    comments.append("A3: Tỷ số khả năng thanh toán tức thời ={:.0f}%. Score= {:.0f}/10.".format(round(100*company.A3_cash_ratio,1),s2))
    comments.append(remarkA)
    comments.append("B. Khả năng cân đối vốn: Score= {:.0f}/30.".format(s3+s4+s5))

    comments.append("B1: Tỷ số nợ trên tổng tài sản ={:.0f}%. Score= {:.0f}/10.".format(round(100*company.B1_debt_ratio,1),s3))
    comments.append("B2: Tỷ số Nợ phải trả trên Vốn chủ sở hữu ={:.0f}%. Score= {:.0f}/10.".format(round(100*company.B2_debt_to_equity_ratio,1),s4))
    comments.append("B3: Tỷ số khả năng thanh toán lãi vay ={:.0f}%. Score= {:.0f}/10.".format(round(100*company.B3_interest_coverage_ratio,1),s5))
    comments.append(remarkB)
    comments.append("C. Hiệu quả hoạt động: Score= {:.0f}/30.".format(s6+s7+s8))

    comments.append("C1: Vòng quay hàng tồn kho ={:.0f}%. Score= {:.0f}/10.".format(round(100*company.C1_inventory_turnover,1),s6))
    comments.append("C2: Kỳ thu tiền trung bình ={:.0f}%. Score= {:.0f}/10.".format(round(100*company.C2_average_collection_period,1),s7))
    comments.append("C3: Vòng quay tài sản cố định ={:.0f}%. Score= {:.0f}/10.".format(round(100*company.C3_fixed_asset_turnover,1),s8))
    comments.append(remarkC)
    comments.append("D. Khả năng sinh lợi: Score= {:.0f}/30.".format(s9+s10+s11))

    comments.append("D1: Tỷ suất doanh lợi doanh thu (ROS) ={:.0f}%. Score= {:.0f}/10.".format(round(company.D1_revenue_ratio,1),s9))
    comments.append("D2: Tỷ số khả năng sinh lời cơ bản của tài sản ={:.0f}%. Score= {:.0f}/10.".format(round(company.D2_basic_profitability_ratio,1),s10))
    comments.append("D3: Tỷ suất doanh lợi tổng tài sản (ROA) ={:.0f}%. Score= {:.0f}/10.".format(round(company.D3_return_on_assets,1),s11))
    comments.append(remarkD)

    return comments

