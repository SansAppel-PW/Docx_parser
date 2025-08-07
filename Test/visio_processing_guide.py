#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Visio 嵌入流程图处理指南

这是一个针对从Word文档中提取的Visio嵌入对象（EMF格式）的专门处理方案
"""

import json
from pathlib import Path

def create_visio_processing_guide():
    """创建Visio处理指南"""
    
    guide = {
        "visio_emf_processing_guide": {
            "overview": "Visio嵌入流程图的综合处理策略",
            "file_characteristics": {
                "format": "Enhanced Metafile (EMF)",
                "source": "Microsoft Visio嵌入对象",
                "typical_size": "50KB - 2MB",
                "complexity": "包含矢量图形、文本、连接线等复杂元素",
                "advantages": [
                    "保持矢量图形质量",
                    "包含完整的绘图信息",
                    "文件相对较小"
                ],
                "limitations": [
                    "不能直接在浏览器中查看",
                    "自动转换工具支持有限",
                    "包含的文本信息不易提取"
                ]
            },
            
            "processing_strategies": {
                "strategy_1_preserve_original": {
                    "name": "保留原始策略（推荐用于RAG）",
                    "description": "保留EMF文件作为原始参考，同时创建处理指导信息",
                    "advantages": [
                        "保持原始信息完整性",
                        "避免转换过程中的信息丢失",
                        "为用户提供多种查看选项"
                    ],
                    "use_cases": [
                        "RAG系统知识库构建",
                        "文档归档和管理",
                        "需要保持原始格式的场景"
                    ],
                    "implementation": {
                        "keep_emf": "保留原始EMF文件",
                        "create_metadata": "生成详细的元数据信息",
                        "provide_guidance": "提供转换和查看指导",
                        "extract_text": "尽可能提取文本信息用于搜索"
                    }
                },
                
                "strategy_2_hybrid_conversion": {
                    "name": "混合转换策略",
                    "description": "尝试自动转换，失败时保留原始文件并提供手动转换指导",
                    "advantages": [
                        "自动化处理简单图形",
                        "为复杂图形提供备选方案",
                        "用户体验友好"
                    ],
                    "conversion_priority": [
                        "1. 尝试系统工具转换（sips, qlmanage）",
                        "2. 尝试专业软件转换（ImageMagick, Inkscape）",
                        "3. 生成转换指导信息",
                        "4. 保留原始文件作为备份"
                    ]
                },
                
                "strategy_3_manual_conversion": {
                    "name": "手动转换策略",
                    "description": "提供详细的手动转换指导，适用于需要高质量输出的场景",
                    "tools": [
                        {
                            "name": "Microsoft Visio",
                            "steps": [
                                "双击EMF文件，用Visio打开",
                                "选择 文件 > 导出 > 更改文件类型",
                                "选择PNG或SVG格式",
                                "设置适当的分辨率（建议300 DPI）",
                                "导出文件"
                            ],
                            "quality": "最高",
                            "requirements": "需要安装Visio"
                        },
                        {
                            "name": "LibreOffice Draw",
                            "steps": [
                                "打开LibreOffice Draw",
                                "选择 文件 > 打开",
                                "选择EMF文件",
                                "选择 文件 > 导出为 > 导出为图像",
                                "选择PNG格式并设置质量",
                                "保存文件"
                            ],
                            "quality": "高",
                            "requirements": "免费软件"
                        },
                        {
                            "name": "在线转换工具",
                            "steps": [
                                "访问 https://convertio.co/emf-png/",
                                "上传EMF文件",
                                "选择输出格式（PNG推荐）",
                                "下载转换结果"
                            ],
                            "quality": "中等",
                            "requirements": "网络连接"
                        }
                    ]
                }
            },
            
            "rag_integration_recommendations": {
                "text_extraction": {
                    "description": "从Visio流程图中提取文本信息用于RAG检索",
                    "methods": [
                        "解析嵌入对象的元数据",
                        "使用OCR技术从转换后的图像中提取文本",
                        "分析文档上下文获取相关描述"
                    ],
                    "challenges": [
                        "EMF格式的文本提取困难",
                        "流程图布局信息缺失",
                        "连接关系无法直接获取"
                    ]
                },
                
                "metadata_enhancement": {
                    "description": "增强元数据以提高RAG系统的理解能力",
                    "suggested_fields": [
                        "diagram_type: 流程图类型（业务流程、技术流程等）",
                        "complexity_level: 复杂度评估",
                        "estimated_nodes: 估算的节点数量", 
                        "domain: 业务领域（如：PLM、制造、研发）",
                        "keywords: 从文件名和上下文提取的关键词"
                    ]
                },
                
                "retrieval_strategy": {
                    "description": "RAG系统中的检索策略",
                    "primary_search": "基于文档上下文和标题进行文本检索",
                    "fallback_search": "基于元数据和关键词进行模糊匹配",
                    "visual_search": "如果转换成功，可以使用图像相似性搜索",
                    "contextual_expansion": "利用文档中的相关段落扩展上下文"
                }
            },
            
            "implementation_examples": {
                "basic_processing": {
                    "description": "基础处理示例",
                    "steps": [
                        "检测Visio嵌入对象",
                        "提取EMF文件", 
                        "分析文件大小和复杂度",
                        "尝试自动转换",
                        "生成处理报告",
                        "创建RAG友好的元数据"
                    ]
                },
                
                "advanced_processing": {
                    "description": "高级处理示例",
                    "steps": [
                        "多格式转换尝试",
                        "文本提取和OCR",
                        "图像分析和特征提取",
                        "上下文关联分析",
                        "生成语义标签",
                        "创建检索索引"
                    ]
                }
            },
            
            "quality_assessment": {
                "success_indicators": [
                    "EMF文件完整保存",
                    "转换指导信息完备",
                    "元数据信息丰富",
                    "RAG系统能够检索到相关内容"
                ],
                "failure_indicators": [
                    "EMF文件损坏或丢失",
                    "缺少处理指导信息",
                    "元数据不完整",
                    "RAG系统无法检索"
                ]
            }
        }
    }
    
    return guide

def generate_visio_processing_report(emf_file_path, document_context=None):
    """
    为特定的Visio EMF文件生成处理报告
    
    Args:
        emf_file_path (str): EMF文件路径
        document_context (dict): 文档上下文信息
        
    Returns:
        dict: 处理报告
    """
    
    emf_path = Path(emf_file_path)
    file_stat = emf_path.stat() if emf_path.exists() else None
    
    report = {
        "file_info": {
            "path": str(emf_file_path),
            "exists": emf_path.exists(),
            "size_kb": round(file_stat.st_size / 1024, 2) if file_stat else 0,
            "complexity": "unknown"
        },
        
        "processing_status": {
            "extracted": emf_path.exists(),
            "conversion_attempted": False,
            "conversion_successful": False,
            "manual_guidance_provided": True
        },
        
        "rag_integration": {
            "searchable_content": [],
            "metadata_fields": {},
            "retrieval_tags": []
        },
        
        "recommendations": [
            "保留原始EMF文件用于高质量查看",
            "使用专业软件进行手动转换",
            "在RAG系统中基于上下文进行检索",
            "考虑使用多模态模型分析转换后的图像"
        ]
    }
    
    # 如果有文档上下文，增强报告
    if document_context:
        # 提取可搜索的内容
        if 'title' in document_context:
            report['rag_integration']['searchable_content'].append(document_context['title'])
            report['rag_integration']['retrieval_tags'].append('title:' + document_context['title'])
        
        if 'section' in document_context:
            report['rag_integration']['searchable_content'].append(document_context['section'])
            report['rag_integration']['retrieval_tags'].append('section:' + document_context['section'])
        
        # 添加元数据
        report['rag_integration']['metadata_fields'].update({
            'diagram_type': '业务流程图',
            'source_application': 'Microsoft Visio',
            'extraction_method': 'embedded_object_parser',
            'format': 'Enhanced Metafile (EMF)'
        })
    
    # 根据文件大小估算复杂度
    if file_stat:
        if file_stat.st_size < 50 * 1024:  # < 50KB
            report['file_info']['complexity'] = 'simple'
            report['recommendations'].insert(0, "文件较小，可尝试自动转换工具")
        elif file_stat.st_size < 200 * 1024:  # < 200KB  
            report['file_info']['complexity'] = 'medium'
            report['recommendations'].insert(0, "中等复杂度，建议使用专业软件转换")
        else:
            report['file_info']['complexity'] = 'complex'
            report['recommendations'].insert(0, "复杂流程图，强烈建议使用Visio或LibreOffice Draw处理")
    
    return report

def main():
    """主函数 - 生成Visio处理指南文档"""
    
    # 生成处理指南
    guide = create_visio_processing_guide()
    
    # 保存到文件
    guide_file = Path('visio_emf_processing_guide.json')
    with open(guide_file, 'w', encoding='utf-8') as f:
        json.dump(guide, f, indent=2, ensure_ascii=False)
    
    print(f"✅ Visio EMF处理指南已生成: {guide_file}")
    
    # 生成示例报告
    sample_context = {
        'title': '电子机电料3D建模流程',
        'section': '流程示意图',
        'description': 'Visio绘图展示了完整的3D建模业务流程'
    }
    
    sample_report = generate_visio_processing_report(
        '/path/to/sample/embedded_obj_visio.emf',
        sample_context
    )
    
    # 保存示例报告
    sample_file = Path('visio_processing_sample_report.json')
    with open(sample_file, 'w', encoding='utf-8') as f:
        json.dump(sample_report, f, indent=2, ensure_ascii=False)
    
    print(f"✅ 示例处理报告已生成: {sample_file}")
    
    # 打印处理策略摘要
    print("\n📋 Visio EMF处理策略摘要:")
    print("=" * 50)
    
    strategies = guide['visio_emf_processing_guide']['processing_strategies']
    for key, strategy in strategies.items():
        print(f"\n{strategy['name']}:")
        print(f"  描述: {strategy['description']}")
        if 'use_cases' in strategy:
            print(f"  适用场景: {', '.join(strategy['use_cases'])}")
    
    print(f"\n🔍 RAG集成建议:")
    rag_rec = guide['visio_emf_processing_guide']['rag_integration_recommendations']
    print(f"  文本提取: {rag_rec['text_extraction']['description']}")
    print(f"  检索策略: {rag_rec['retrieval_strategy']['description']}")
    
    return guide, sample_report

if __name__ == '__main__':
    main()
