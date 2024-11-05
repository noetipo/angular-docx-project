// word-document.service.ts
import {Component} from '@angular/core';
import {DocumentService} from './word-documen.service';

@Component({
  selector: 'app-document',
  standalone: true,
  templateUrl: './document.component.html',
  styleUrl: './document.component.scss'
})
export class DocumentComponent {
  constructor(private documentService: DocumentService) {
  }

  downloadDocument(): void {
    this.documentService.generateDocument();
  }
}
