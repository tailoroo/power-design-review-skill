#!/usr/bin/env python3
"""
合并去重脚本 - 合并多个角色的检查结果，执行精确去重

使用方法:
    python merge_results.py <input_json_files...> -o <output_file>

示例:
    python merge_results.py general-checker.json senior-engineer.json reviewer.json -o merged_issues.json
"""

import json
import argparse
from pathlib import Path
from typing import List, Dict, Any
from collections import defaultdict
import hashlib


def load_json(file_path: str) -> Dict[str, Any]:
    """加载JSON文件"""
    with open(file_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def save_json(data: Dict[str, Any], file_path: str) -> None:
    """保存JSON文件"""
    with open(file_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def generate_issue_hash(issue: Dict[str, Any]) -> str:
    """生成问题的唯一哈希值（基于位置和描述）"""
    location = issue.get('location', '').strip()
    description = issue.get('description', '').strip()
    category = issue.get('category', '').strip()

    # 使用位置+描述+分类生成哈希
    key = f"{location}|{description}|{category}"
    return hashlib.md5(key.encode('utf-8')).hexdigest()


def merge_issues(all_results: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """合并所有问题，执行去重"""
    issue_map = {}  # hash -> issue
    issue_sources = defaultdict(list)  # hash -> [sources]

    for result in all_results:
        agent = result.get('agent', 'unknown')
        issues = result.get('issues', [])

        for issue in issues:
            issue_hash = generate_issue_hash(issue)

            if issue_hash not in issue_map:
                issue_map[issue_hash] = issue.copy()
                issue_map[issue_hash]['sources'] = [agent]
            else:
                # 合并来源
                issue_map[issue_hash]['sources'].append(agent)
                # 合并位置（如果有多个位置）
                existing_locations = issue_map[issue_hash].get('evidence', {}).get('locations', [])
                new_locations = issue.get('evidence', {}).get('locations', [])
                merged_locations = list(set(existing_locations + new_locations))
                if merged_locations:
                    if 'evidence' not in issue_map[issue_hash]:
                        issue_map[issue_hash]['evidence'] = {}
                    issue_map[issue_hash]['evidence']['locations'] = merged_locations
                    # 更新位置描述
                    if len(merged_locations) > 1:
                        issue_map[issue_hash]['location'] = '、'.join(merged_locations[:3])
                        if len(merged_locations) > 3:
                            issue_map[issue_hash]['location'] += f'等{len(merged_locations)}处'

    return list(issue_map.values())


def apply_review_feedback(merged_issues: List[Dict[str, Any]],
                          review_result: Dict[str, Any]) -> List[Dict[str, Any]]:
    """应用审查员的反馈"""
    # 获取无效问题ID
    invalid_ids = set()
    for invalid in review_result.get('invalid_issues', []):
        invalid_ids.add(invalid.get('id'))

    # 获取遗漏问题
    missed_issues = review_result.get('missed_issues', [])

    # 过滤无效问题
    filtered_issues = [
        issue for issue in merged_issues
        if issue.get('id') not in invalid_ids
    ]

    # 添加遗漏问题
    for missed in missed_issues:
        missed_issue = {
            'id': missed.get('id'),
            'category': missed.get('category'),
            'description': missed.get('description'),
            'location': missed.get('location'),
            'suggestion': missed.get('suggestion'),
            'severity': missed.get('severity', 'medium'),
            'source': 'reviewer_discovery'
        }
        filtered_issues.append(missed_issue)

    return filtered_issues


def apply_corrections(merged_issues: List[Dict[str, Any]],
                      review_result: Dict[str, Any]) -> List[Dict[str, Any]]:
    """应用审查员的修正建议"""
    corrections = {
        c.get('id'): c
        for c in review_result.get('correction_suggestions', [])
    }

    for issue in merged_issues:
        issue_id = issue.get('id')
        if issue_id in corrections:
            correction = corrections[issue_id]
            field = correction.get('field')
            suggested = correction.get('suggested')
            if field and suggested:
                issue[field] = suggested
                issue['corrected'] = True
                issue['correction_reason'] = correction.get('reason')

    return merged_issues


def generate_summary(issues: List[Dict[str, Any]]) -> Dict[str, Any]:
    """生成问题统计摘要"""
    by_category = defaultdict(int)
    by_severity = defaultdict(int)
    by_source = defaultdict(int)

    for issue in issues:
        by_category[issue.get('category', '其他')] += 1
        by_severity[issue.get('severity', 'medium')] += 1
        for source in issue.get('sources', [issue.get('source', 'unknown')]):
            by_source[source] += 1

    return {
        'total_issues': len(issues),
        'by_category': dict(by_category),
        'by_severity': dict(by_severity),
        'by_source': dict(by_source)
    }


def renumber_issues(issues: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """重新编号问题"""
    for i, issue in enumerate(issues, 1):
        issue['序号'] = i
    return issues


def convert_to_table_format(issues: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """转换为5列表格格式"""
    table_rows = []
    for issue in issues:
        table_rows.append({
            '序号': issue.get('序号', 0),
            '分类': issue.get('category', '其他'),
            '问题详细内容': issue.get('description', ''),
            '问题位置': issue.get('location', ''),
            '修改建议': issue.get('suggestion', '')
        })
    return table_rows


def main():
    parser = argparse.ArgumentParser(description='合并去重检查结果')
    parser.add_argument('input_files', nargs='+', help='输入JSON文件')
    parser.add_argument('-o', '--output', required=True, help='输出文件路径')
    parser.add_argument('--format', choices=['json', 'table'], default='json',
                        help='输出格式：json或table')
    parser.add_argument('--review', help='审查结果JSON文件（可选）')

    args = parser.parse_args()

    # 加载所有结果文件
    all_results = []
    for file_path in args.input_files:
        try:
            result = load_json(file_path)
            all_results.append(result)
            print(f"已加载: {file_path}")
        except Exception as e:
            print(f"警告: 无法加载 {file_path}: {e}")

    # 合并问题
    merged_issues = merge_issues(all_results)
    print(f"合并后问题数: {len(merged_issues)}")

    # 应用审查反馈（如果有）
    if args.review:
        try:
            review_result = load_json(args.review)
            merged_issues = apply_review_feedback(merged_issues, review_result)
            merged_issues = apply_corrections(merged_issues, review_result)
            print(f"应用审查反馈后问题数: {len(merged_issues)}")
        except Exception as e:
            print(f"警告: 无法加载审查结果 {args.review}: {e}")

    # 重新编号
    merged_issues = renumber_issues(merged_issues)

    # 生成输出
    if args.format == 'table':
        output_data = {
            'issues': convert_to_table_format(merged_issues),
            'summary': generate_summary(merged_issues)
        }
    else:
        output_data = {
            'merged_issues': merged_issues,
            'summary': generate_summary(merged_issues),
            'sources': [r.get('agent', 'unknown') for r in all_results]
        }

    # 保存输出
    save_json(output_data, args.output)
    print(f"已保存到: {args.output}")


def generate_summary(issues: List[Dict[str, Any]]) -> Dict[str, Any]:
    """生成摘要统计"""
    by_category = defaultdict(int)
    by_severity = defaultdict(int)

    for issue in issues:
        by_category[issue.get('category', '其他')] += 1
        by_severity[issue.get('severity', 'medium')] += 1

    return {
        'total_issues': len(issues),
        'by_category': dict(by_category),
        'by_severity': dict(by_severity)
    }


if __name__ == '__main__':
    main()
